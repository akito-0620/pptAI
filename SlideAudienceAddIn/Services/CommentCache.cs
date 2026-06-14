using System.Collections.Concurrent;
using System.Collections.Generic;
using SlideAudienceAddIn.Models;

namespace SlideAudienceAddIn.Services
{
    public class CommentCache
    {
        private readonly ConcurrentDictionary<int, IReadOnlyList<AudienceComment>> _commentsBySlideId =
            new ConcurrentDictionary<int, IReadOnlyList<AudienceComment>>();

        public bool TryGet(int slideId, out IReadOnlyList<AudienceComment> comments)
        {
            return _commentsBySlideId.TryGetValue(slideId, out comments);
        }

        public void Set(int slideId, IReadOnlyList<AudienceComment> comments)
        {
            _commentsBySlideId[slideId] = comments;
        }

        public void Clear()
        {
            _commentsBySlideId.Clear();
        }
    }
}
