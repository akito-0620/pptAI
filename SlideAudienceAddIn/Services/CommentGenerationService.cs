using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using SlideAudienceAddIn.Models;
using SlideAudienceAddIn.Utils;

namespace SlideAudienceAddIn.Services
{
    public class CommentGenerationService
    {
        private readonly AppSettings _settings;
        private readonly HttpClient _httpClient;

        public CommentGenerationService(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(Math.Max(5, _settings.Gemini.TimeoutSeconds))
            };
        }

        public async Task<CommentGenerationResult> GenerateForSlideAsync(
            int slideId,
            int slideIndex,
            string slideImagePath,
            string extractedText,
            CancellationToken cancellationToken)
        {
            var stopwatch = Stopwatch.StartNew();
            var promptLength = SafePromptLength(slideId, slideIndex, extractedText);
            var extractedTextLength = string.IsNullOrEmpty(extractedText) ? 0 : extractedText.Length;

            if (!ShouldUseApi(out var apiKey, out var fallbackReason))
            {
                Debug.WriteLine(
                    $"[SlideAudience] CommentGenerationService: EnableApi=false or API key missing; returning dummy comments. slideId={slideId}, slideIndex={slideIndex}");
                var dummyComments = DummyComments(slideIndex, extractedText);
                Debug.WriteLine($"[SlideAudience] CommentGenerationService: dummy comments count={dummyComments.Count}");
                foreach (var comment in dummyComments)
                {
                    Debug.WriteLine($"[SlideAudience] CommentGenerationService dummy comment.Text: {comment.Text}, persona={comment.Persona}");
                }

                return Finish(dummyComments, stopwatch, false, fallbackReason, promptLength, extractedTextLength);
            }

            try
            {
                Debug.WriteLine(
                    $"[SlideAudience] CommentGenerationService: Gemini API call start. slideId={slideId}, slideIndex={slideIndex}, imagePath={slideImagePath}");
                var comments = await GenerateWithGeminiAsync(
                    apiKey,
                    slideId,
                    slideIndex,
                    slideImagePath,
                    extractedText,
                    cancellationToken);

                Debug.WriteLine(
                    $"[SlideAudience] CommentGenerationService: Gemini API call completed. comments count={comments?.Count ?? 0}, latencyMs={stopwatch.ElapsedMilliseconds}");
                foreach (var comment in comments ?? new List<AudienceComment>())
                {
                    Debug.WriteLine($"[SlideAudience] CommentGenerationService Gemini comment.Text: {comment.Text}, persona={comment.Persona}");
                }

                return Finish(comments, stopwatch, true, null, promptLength, extractedTextLength);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] CommentGenerationService: Gemini API failed; falling back to dummy comments");
                Debug.WriteLine($"[SlideAudience] CommentGenerationService fallback reason: {ex.Message}");
                Debug.WriteLine(ex.ToString());

                var dummyComments = DummyComments(slideIndex, extractedText);
                Debug.WriteLine($"[SlideAudience] CommentGenerationService fallback dummy comments count={dummyComments.Count}");
                foreach (var comment in dummyComments)
                {
                    Debug.WriteLine($"[SlideAudience] CommentGenerationService fallback dummy comment.Text: {comment.Text}, persona={comment.Persona}");
                }

                return new CommentGenerationResult
                {
                    Comments = dummyComments,
                    LatencyMs = stopwatch.ElapsedMilliseconds,
                    UsedApi = false,
                    ErrorMessage = ex.Message,
                    FallbackReason = ex.Message,
                    Model = _settings.Gemini.Model,
                    PromptLength = promptLength,
                    ExtractedTextLength = extractedTextLength
                };
            }
        }

        private bool ShouldUseApi(out string apiKey, out string fallbackReason)
        {
            apiKey = null;
            fallbackReason = null;
            if (!_settings.Gemini.EnableApi)
            {
                Debug.WriteLine("[SlideAudience] CommentGenerationService: Gemini EnableApi=false; using local dummy comments");
                fallbackReason = "EnableApi=false";
                return false;
            }

            apiKey = Environment.GetEnvironmentVariable(_settings.Gemini.ApiKeyEnvironmentVariable);
            apiKey = string.IsNullOrWhiteSpace(apiKey) ? null : apiKey.Trim();
            var hasApiKey = !string.IsNullOrWhiteSpace(apiKey);
            if (!hasApiKey)
            {
                Debug.WriteLine(
                    $"[SlideAudience] CommentGenerationService: Gemini API key is missing in environment variable '{_settings.Gemini.ApiKeyEnvironmentVariable}'; using local dummy comments");
                fallbackReason = "API key missing";
            }

            return hasApiKey;
        }

        private async Task<IReadOnlyList<AudienceComment>> GenerateWithGeminiAsync(
            string apiKey,
            int slideId,
            int slideIndex,
            string slideImagePath,
            string extractedText,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(slideImagePath) || !File.Exists(slideImagePath))
            {
                throw new FileNotFoundException("Slide image export was not found.", slideImagePath);
            }

            Debug.WriteLine($"[SlideAudience] CommentGenerationService: reading slide image bytes from {slideImagePath}");
            var imageBase64 = Convert.ToBase64String(File.ReadAllBytes(slideImagePath));
            var prompt = BuildPrompt(slideId, slideIndex, extractedText);
            Debug.WriteLine($"[SlideAudience] CommentGenerationService: prompt length={prompt.Length}, extractedText length={(extractedText ?? string.Empty).Length}");

            var request = new Dictionary<string, object>
            {
                ["contents"] = new object[]
                {
                    new Dictionary<string, object>
                    {
                        ["role"] = "user",
                        ["parts"] = new object[]
                        {
                            new Dictionary<string, object> { ["text"] = prompt },
                            new Dictionary<string, object>
                            {
                                ["inline_data"] = new Dictionary<string, object>
                                {
                                    ["mime_type"] = "image/png",
                                    ["data"] = imageBase64
                                }
                            }
                        }
                    }
                },
                ["generationConfig"] = new Dictionary<string, object>
                {
                    ["temperature"] = 0.78,
                    ["responseMimeType"] = "application/json"
                }
            };

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/{Uri.EscapeDataString(_settings.Gemini.Model)}:generateContent";
            using (var content = new StringContent(JsonHelper.Serialize(request), Encoding.UTF8, "application/json"))
            using (var requestMessage = new HttpRequestMessage(HttpMethod.Post, url))
            {
                requestMessage.Headers.Add("x-goog-api-key", apiKey);
                requestMessage.Content = content;
                Debug.WriteLine($"[SlideAudience] CommentGenerationService: sending Gemini request. model={_settings.Gemini.Model}");
                using (var response = await _httpClient.SendAsync(requestMessage, cancellationToken))
                {
                    Debug.WriteLine($"[SlideAudience] CommentGenerationService: Gemini response status={(int)response.StatusCode} {response.ReasonPhrase}");
                    var responseText = await response.Content.ReadAsStringAsync();
                    cancellationToken.ThrowIfCancellationRequested();
                    Debug.WriteLine($"[SlideAudience] CommentGenerationService: Gemini response body length={responseText.Length}");

                    if (!response.IsSuccessStatusCode)
                    {
                        Debug.WriteLine("[SlideAudience] CommentGenerationService: Gemini error response body:");
                        Debug.WriteLine(responseText);
                        response.EnsureSuccessStatusCode();
                    }

                    var jsonText = ExtractGeminiText(responseText);
                    Debug.WriteLine("[SlideAudience] CommentGenerationService: Gemini extracted JSON text:");
                    Debug.WriteLine(jsonText);

                    var parsedComments = ParseCommentJson(jsonText);
                    var normalizedComments = NormalizeComments(parsedComments, fallbackSlideIndex: slideIndex, fallbackText: extractedText);
                    if (normalizedComments == null || normalizedComments.Count == 0)
                    {
                        throw new InvalidOperationException("Gemini comments were empty after normalization.");
                    }

                    return normalizedComments;
                }
            }
        }

        private string BuildPrompt(int slideId, int slideIndex, string extractedText)
        {
            var count = Math.Max(1, _settings.Comments.MaxCommentsPerSlide);
            return
$@"あなたはプレゼンテーションを見ている観客です。
入力されたスライド画像とスライド内テキストを見て、観客が一瞬で読める短い一言リアクションを日本語で{count}件生成してください。

方針:
- 発表者の代弁ではなく観客の内心として書く
- 説明文ではなく思わず出る短い反応にする
- 攻撃的にしすぎない
- スライドに書かれていない内容を断定しすぎない
- 句点「。」は使わない
- JSONのみで返す

コメント種別:
- understanding: 理解確認
- interest: 興味や共感
- question: 疑問や批判

ペルソナ:
- beginner
- expert
- skeptic
- curious
- practical
- research_evaluator
- designer
- experienced_speaker
- tsukkomi
- empathetic

制約:
- コメントは必ず{count}件
- 1コメントは10から15文字程度
- 最大18文字以内
- typeは understanding, interest, question のいずれか
- personaは上記ペルソナから選ぶ
- 同じ言い回しを避ける

モード: {_settings.Comments.Mode}
スライド番号: {slideIndex}
SlideID: {slideId}

スライド内テキスト:
{(string.IsNullOrWhiteSpace(extractedText) ? "(テキスト抽出なし)" : extractedText)}

返答形式:
{{
  ""comments"": [
    {{""type"": ""understanding"", ""persona"": ""beginner"", ""text"": ""つまり要点は？""}},
    {{""type"": ""interest"", ""persona"": ""curious"", ""text"": ""雰囲気づくり大事""}},
    {{""type"": ""question"", ""persona"": ""skeptic"", ""text"": ""なぜその表現？""}}
  ]
}}";
        }

        private static string ExtractGeminiText(string responseText)
        {
            var root = JsonHelper.Deserialize<Dictionary<string, object>>(responseText);
            var candidates = AsItems(root["candidates"]);
            var firstCandidate = candidates.FirstOrDefault() as Dictionary<string, object>;
            var content = firstCandidate?["content"] as Dictionary<string, object>;
            var parts = content != null && content.ContainsKey("parts") ? AsItems(content["parts"]) : new List<object>();
            var firstPart = parts.FirstOrDefault() as Dictionary<string, object>;
            var text = firstPart?["text"] as string;

            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("Gemini response did not include text.");
            }

            return StripJsonFence(text.Trim());
        }

        private IReadOnlyList<AudienceComment> ParseCommentJson(string jsonText)
        {
            var root = JsonHelper.Deserialize<Dictionary<string, object>>(jsonText);
            var comments = new List<AudienceComment>();
            if (!root.TryGetValue("comments", out var rawComments))
            {
                throw new InvalidOperationException("Generated JSON did not include comments.");
            }

            foreach (var rawItem in AsItems(rawComments))
            {
                var item = rawItem as Dictionary<string, object>;
                if (item == null)
                {
                    continue;
                }

                var type = item.TryGetValue("type", out var rawType) ? Convert.ToString(rawType) : "comment";
                var persona = item.TryGetValue("persona", out var rawPersona) ? Convert.ToString(rawPersona) : null;
                var text = item.TryGetValue("text", out var rawText) ? Convert.ToString(rawText) : string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    comments.Add(AudienceComment.Create(type, text.Trim(), persona: persona));
                }
            }

            if (comments.Count == 0)
            {
                throw new InvalidOperationException("Generated JSON comments were empty.");
            }

            return comments;
        }

        private static List<object> AsItems(object value)
        {
            if (value is ArrayList arrayList)
            {
                return arrayList.Cast<object>().ToList();
            }

            if (value is object[] array)
            {
                return array.Cast<object>().ToList();
            }

            if (value is IEnumerable enumerable && !(value is string))
            {
                return enumerable.Cast<object>().ToList();
            }

            return new List<object>();
        }

        private IReadOnlyList<AudienceComment> DummyComments(int slideIndex, string extractedText)
        {
            var topic = FirstUsefulLine(extractedText);
            var comments = new List<AudienceComment>
            {
                AudienceComment.Create("understanding", string.IsNullOrWhiteSpace(topic) ? "つまり要点は？" : $"つまり{TrimForComment(topic, 8)}？", persona: "beginner"),
                AudienceComment.Create("interest", "ちょっと気になる", persona: "curious"),
                AudienceComment.Create("question", "なぜその表現？", persona: "skeptic"),
                AudienceComment.Create("interest", "使い道ありそう", persona: "practical"),
                AudienceComment.Create("question", "評価方法はどこ？", persona: "research_evaluator"),
                AudienceComment.Create("interest", "見せ方が効いてる", persona: "designer"),
                AudienceComment.Create("understanding", "流れは追いやすい", persona: "experienced_speaker"),
                AudienceComment.Create("question", "そこ深掘りしたい", persona: "expert"),
                AudienceComment.Create("interest", "ここ好きかも", persona: "empathetic"),
                AudienceComment.Create("question", "今の一言強いな", persona: "tsukkomi")
            };

            return NormalizeComments(comments, slideIndex, extractedText);
        }

        private IReadOnlyList<AudienceComment> NormalizeComments(
            IEnumerable<AudienceComment> comments,
            int fallbackSlideIndex,
            string fallbackText)
        {
            var allowedTypes = RequestedTypes().ToList();
            var normalized = comments
                .Where(comment => comment != null)
                .Where(comment => !string.IsNullOrWhiteSpace(comment.Text))
                .Select(comment =>
                {
                    var type = string.IsNullOrWhiteSpace(comment.Type) ? allowedTypes[0] : comment.Type.Trim();
                    return AudienceComment.Create(
                        type,
                        NormalizeCommentText(comment.Text, type),
                        comment.Confidence,
                        NormalizePersona(comment.Persona));
                })
                .Where(comment => allowedTypes.Contains(comment.Type))
                .Take(Math.Max(1, _settings.Comments.MaxCommentsPerSlide))
                .ToList();

            if (normalized.Count > 0)
            {
                return normalized;
            }

            return DummyComments(fallbackSlideIndex, fallbackText);
        }

        private IEnumerable<string> RequestedTypes()
        {
            switch ((_settings.Comments.Mode ?? "Mixed").Trim())
            {
                case "UnderstandingOnly":
                    return new[] { "understanding" };
                case "InterestOnly":
                    return new[] { "interest" };
                case "CriticalOnly":
                    return new[] { "question" };
                default:
                    return new[] { "understanding", "interest", "question" };
            }
        }

        private static string StripJsonFence(string text)
        {
            if (text.StartsWith("```", StringComparison.Ordinal))
            {
                var firstNewLine = text.IndexOf('\n');
                var lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
                if (firstNewLine >= 0 && lastFence > firstNewLine)
                {
                    return text.Substring(firstNewLine + 1, lastFence - firstNewLine - 1).Trim();
                }
            }

            return text;
        }

        private static string FirstUsefulLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            return text
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim())
                .FirstOrDefault(line => line.Length >= 2);
        }

        private static string TrimForComment(string text, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(text) || text.Length <= maxLength)
            {
                return text;
            }

            if (maxLength <= 3)
            {
                return text.Substring(0, maxLength);
            }

            return text.Substring(0, maxLength - 1) + "…";
        }

        private static string NormalizeCommentText(string text, string type)
        {
            var normalized = (text ?? string.Empty)
                .Replace("。", string.Empty)
                .Trim();
            var before = normalized;
            if (normalized.Length < 5 || normalized.Length > 18)
            {
                normalized = ReplacementComment(type);
            }

            normalized = TrimForComment(normalized.Replace("。", string.Empty).Trim(), 18);
            Debug.WriteLine($"[SlideAudience] comment length normalized: before='{before}'({before.Length}), after='{normalized}'({normalized.Length})");
            return normalized;
        }

        private static string ReplacementComment(string type)
        {
            switch ((type ?? string.Empty).Trim())
            {
                case "understanding":
                    return "つまり要点は？";
                case "interest":
                    return "ちょっと気になる";
                case "question":
                    return "なぜその表現？";
                default:
                    return "観客反応ありそう";
            }
        }

        private static string NormalizePersona(string persona)
        {
            return string.IsNullOrWhiteSpace(persona) ? "audience" : persona.Trim();
        }

        private CommentGenerationResult Finish(
            IReadOnlyList<AudienceComment> comments,
            Stopwatch stopwatch,
            bool usedApi,
            string fallbackReason,
            int promptLength,
            int extractedTextLength)
        {
            Debug.WriteLine(
                $"[SlideAudience] CommentGenerationService result: usedApi={usedApi}, latencyMs={stopwatch.ElapsedMilliseconds}, comments count={comments?.Count ?? 0}");

            return new CommentGenerationResult
            {
                Comments = comments,
                LatencyMs = stopwatch.ElapsedMilliseconds,
                UsedApi = usedApi,
                FallbackReason = fallbackReason,
                Model = _settings.Gemini.Model,
                PromptLength = promptLength,
                ExtractedTextLength = extractedTextLength
            };
        }

        private int SafePromptLength(int slideId, int slideIndex, string extractedText)
        {
            try
            {
                return BuildPrompt(slideId, slideIndex, extractedText).Length;
            }
            catch
            {
                return 0;
            }
        }
    }
}
