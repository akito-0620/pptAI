# SlideAudience

SlideAudience is a Windows PowerPoint add-in prototype that shows short AI-style audience comments on top of a running slide show.

The first milestone follows the design document's MVP:

- Detect slide show begin / slide change / slide show end.
- Export the current slide to PNG.
- Extract slide text from PowerPoint shapes.
- Show comments in a transparent WPF overlay.
- Cache comments by `SlideID`.
- Save experiment logs as JSONL.
- Edit runtime settings from the PowerPoint ribbon.
- Pregenerate comments for all slides before presenting.
- Fall back to dummy comments when Gemini is disabled or no API key is configured.

## Requirements

- Windows 10 or Windows 11
- Microsoft PowerPoint desktop app
- Visual Studio 2022
- Visual Studio workload: Office/SharePoint development
- Individual component: .NET Framework 4.8 targeting pack

This repository contains the source scaffold and implementation files. The current machine has MSBuild installed, but the VSTO build targets and .NET Framework 4.8 targeting pack are not visible from the command line, so final add-in build and registration should be done in Visual Studio after installing those components.

## Configuration

Copy:

```text
SlideAudienceAddIn/Config/appsettings.example.json
```

to:

```text
SlideAudienceAddIn/Config/appsettings.local.json
```

Then adjust values as needed. Gemini reads the API key from:

```text
GEMINI_API_KEY
```

If the key is missing, or `EnableApi` is false, the add-in displays local dummy comments.

To enable Gemini generation, set:

```json
"EnableApi": true
```

in `appsettings.local.json`, then set the environment variable:

```powershell
setx GEMINI_API_KEY "your_api_key_here"
```

Restart PowerPoint after changing environment variables.

## Run In Visual Studio

1. Open `SlideAudience.sln`.
2. Confirm the project targets a .NET Framework version installed on your machine.
3. Set PowerPoint as the Office debug host if Visual Studio asks.
4. Start debugging.
5. Open a PowerPoint file and start a slide show.
6. Use the `SlideAudience` ribbon tab to enable / disable, generate the current slide, pregenerate all slides, clear cache, or open settings.

## MVP Smoke Test

Use this sequence first, before enabling Gemini:

1. Keep `EnableApi` as `false`.
2. Start debugging from Visual Studio.
3. Start a PowerPoint slide show.
4. Advance to the next slide.
5. Confirm a transparent comment panel appears near the bottom-right of the slide show.
6. End the slide show and confirm the overlay disappears.

If this works, the PowerPoint event handling, slide export, text extraction, overlay, cache, and logging path are all alive.

For lower latency during a real presentation, click `Pregenerate All Slides` before starting the slide show. This fills the SlideID cache so slide changes can reuse generated comments.

## Troubleshooting

- If the project will not load as a VSTO add-in, install the Visual Studio `Office/SharePoint development` workload.
- If `v4.8` targeting fails, install the `.NET Framework 4.8 targeting pack` from the Visual Studio Installer individual components tab.
- If the overlay appears but blocks slide clicks, rebuild after confirming `OverlayWindow.xaml.cs` includes the click-through extended window style setup.
- If Gemini returns an error, switch `EnableApi` back to `false` to continue testing the add-in shell with dummy comments.

## Logs

Experiment logs are written to:

```text
Documents/SlideAudience/logs/
```

Slide exports are written to:

```text
%TEMP%/SlideAudience/exports/
```

By default, logs avoid saving the full slide text for privacy.
