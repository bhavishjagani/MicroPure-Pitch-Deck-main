# MicroPure AI — Pitch Deck

A self-contained, browser-based pitch deck for MicroPure AI by Bhavish Jagani
(9th grade, Bellarmine College Preparatory, San Jose, CA).

## Open it

Just double-click `index.html`, or:

```bash
open index.html
```

No build step, no install. Internet is only used the first time so Chart.js
and the Inter / Space Grotesk fonts can load — after that it works offline.

## Present it

| Action          | Keys                                  |
| --------------- | ------------------------------------- |
| Next slide      | `→`  ·  `Space`  ·  `PageDown`  ·  click background |
| Previous slide  | `←`  ·  `PageUp`                       |
| Jump to slide N | `1` – `9`                             |
| First / Last    | `Home` / `End`                        |
| Save as PDF     | `⌘P` (Cmd-P) → "Save as PDF"          |

For the cleanest export:

1. Open the deck full-screen (`F` in most browsers).
2. `⌘P` → Destination: **Save as PDF**.
3. Layout: **Landscape**, Margins: **None**, Background graphics: **On**.

## Slide map

| #  | Slide                                   |
| -- | --------------------------------------- |
| 01 | Title + tagline                         |
| 02 | The problem                             |
| 03 | The solution + live Analyze screenshot  |
| 04 | Product · Detection Dashboard           |
| 05 | Product Tour · charts / history / settings |
| 06 | 60-second demo video                    |
| 07 | Unique Value Proposition                |
| 08 | Customer Segments                       |
| 09 | Target Markets · TAM / SAM / SOM        |
| 10 | Expected Market Growth                  |
| 11 | Channels · go-to-market                 |
| 12 | Key Activities                          |
| 13 | Key Resources                           |
| 14 | Competitive Advantages                  |
| 15 | User Validation + Testimonials          |
| 16 | Revenue Streams                         |
| 17 | Cost Structure                          |
| 18 | Financial Model · 3-year projection     |
| 19 | Product Roadmap                         |
| 20 | Founder                                 |
| 21 | References                              |
| 22 | Thank you / Q&A                         |

## Screenshots already wired up

These live screenshots are referenced from `assets/` and render automatically:

| File                    | Used on slide(s) |
| ----------------------- | ---------------- |
| `assets/analyze.png`    | 03               |
| `assets/dashboard.png`  | 04               |
| `assets/charts.png`     | 05               |
| `assets/history.png`    | 05               |
| `assets/settings.png`   | 05               |
| `assets/about.png`      | 20               |

## Adding your demo video (slide 06)

Drop a screen-recording into `assets/` (mp4 is best), then in `index.html`
find the section labeled `<!-- 6. DEMO VIDEO -->` and replace the
placeholder `<div class="ph">` with:

```html
<video src="assets/demo-video.mp4" controls autoplay muted loop
       style="width:100%;aspect-ratio:16/9;object-fit:cover;
              border-radius:18px;border:1px solid rgba(148,163,196,.22);
              box-shadow:0 30px 80px rgba(0,0,0,.5);background:#05070f">
</video>
```

The deck is intentionally a single HTML file so it's easy to email, drop on
a USB stick, or open from a Drive link without any setup on the other end.

## Editing copy

All text is in `index.html`. Search for the slide number comment
(`<!-- 4. PRODUCT EVALUATION / SURVEY -->`, etc.) to jump to a section.
Survey numbers, pricing tiers, and reference list are all in plain HTML — no
templating to wrestle with.
