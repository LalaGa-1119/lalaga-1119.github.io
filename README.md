# Lili Gao · Personal Portfolio

> Product Manager · Beijing · 2 years of experience · 6 zero-to-one projects

A single-file personal portfolio that runs without a build step. Its editorial layout is inspired by Oil®, with soft pink tones and vivid gradients.

## Features

- Animated loading screen and line-by-line text reveals
- Grid background, ambient glows, and floating color blocks
- Mascot eyes that follow the pointer or mobile gyroscope
- Reading progress indicator
- Tab navigation and animated statistics
- Project cards with live CSS demos
- Writing section linking to articles published on Woshipm

## Live Site

👉 https://lalaga-1119.github.io/

## Run Locally

Open `index.html` directly in a browser, or run:

```bash
open index.html
```

## Files

The Jiazuo project card embeds the 28-second V5 product introduction. Playback
starts only when the visitor presses play; native controls support seeking,
volume and fullscreen, with inline playback on mobile. A separate video link is
available as a fallback. The original product demo link is unchanged.

| File | Description |
| --- | --- |
| `index.html` | Complete single-file portfolio with inline styles and scripts |
| `project-atelier-v5.mp4` | 1080p H.264/AAC product video, with fast-start metadata for web playback |
| `project-atelier-v5-poster.jpg` | Poster extracted from the same V5 video |
| `article-cover.jpg` | Article thumbnail |
| `profile-photo.jpg` | Portrait used on the homepage |
| `profile-info.xlsx` | Portfolio information in Excel format |
| `generate_excel.py` | Script that generates the spreadsheet; requires `openpyxl` |
