# Rizz Score

A joke web app. You answer ten silly questions about yourself and an LLM rates
your "rizz" out of 100, gives you a rank, a short roast and three tips.

Built for a university proposal by Khant Wai Yan Aung (6703712) and
Kaung San Thu (6703124).

## How it works

1. The browser loads the question list from the server.
2. You pick one option per question.
3. A seven second face scan opens the camera, draws a scanning overlay and asks
   you to pull a silly face. This is optional and there is a skip button.
4. The server builds a prompt from your answers plus whether you pulled a face,
   and sends it to one LLM API.
5. The LLM replies in JSON with a score, a rank title, a roast and three tips.
6. The browser animates the score and shows the card.

## The face scan

The camera feed never leaves the browser. No frame is uploaded, recorded or
saved. All the page does is compare downscaled frames to measure movement, and
the only thing sent to the server is one word: `mocked`, `still` or `skipped`.
If the camera is blocked or missing, the scan is skipped and the app carries on.

If the LLM call fails, the server falls back to a simple local calculation so a
live demo never shows a blank screen.

## Setup

```
npm install
```

Copy `.env.example` to `.env` and fill in your key:

```
LLM_API_KEY=your_key_here
LLM_BASE_URL=https://api.groq.com/openai/v1
LLM_MODEL=openai/gpt-oss-120b
PORT=3000
```

## Run

```
npm start
```

Then open http://localhost:3000

## Files

```
server.js          Express server, question list, prompt, LLM call
public/index.html  Page structure
public/style.css   Styling
public/app.js      Form logic, request, score animation
.env               Your key, never committed
```

## Note

The result is comedy, not an assessment of any real person.

## Proposal slides

`slides/Rizz Score Proposal.pptx` is the deck for class. Open it in PowerPoint
or Keynote. Seven slides, 16:9, with speaker notes on every slide.

`slides/make_pptx.py` rebuilds that file. Edit the script and run
`python3 slides/make_pptx.py` to regenerate it. Needs `pip install python-pptx`.

`slides/proposal.html` is the older twelve slide version as a web page, if you
would rather present from a browser. No server needed.

- Arrow keys, space or click to move between slides
- `N` toggles the speaker notes
- `F` for full screen
- To make a PDF: print the page, choose landscape, and set margins to none

Seven slides, in the order the proposal format asks for:

1. Topic of focus and the team
2. Scope: the functions it can do, and what is cut on purpose
3. Techniques in the browser
4. Techniques on the server, and the reply shape
5. Expected result: the question screen and the face scan
6. Expected result: the score card and the three states
7. Conclusion and references

The deck is drawn from the app's own design: the colours in `public/style.css`,
the pill option chips, the hot to cool gradient, the score meter and the scan
overlay. It uses Inter, Inter Display and JetBrains Mono, so install those
fonts on whatever machine you present from.

`make_pptx.py` measures every line with the real font files before it draws, and
asserts that no two text boxes on a slide overlap and that nothing spills out of
a phone mock or past the bottom of the slide. If an edit makes something too
tall the script fails instead of writing a broken deck.
