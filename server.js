import express from 'express'
import { dirname, join } from 'path'
import { fileURLToPath } from 'url'
import 'dotenv/config'

const __dirname = dirname(fileURLToPath(import.meta.url))

const app = express()
app.use(express.json())
app.use(express.static(join(__dirname, 'public')))

const API_KEY = process.env.LLM_API_KEY
const BASE_URL = process.env.LLM_BASE_URL || 'https://api.groq.com/openai/v1'
const MODEL = process.env.LLM_MODEL || 'openai/gpt-oss-120b'
const PORT = process.env.PORT || 3000

// The questions live here so the form and the prompt can never drift apart.
const QUESTIONS = [
  { id: 'height', label: 'How tall are you?', type: 'select', options: ['Under 160cm', '160 to 170cm', '170 to 180cm', 'Over 180cm'] },
  { id: 'haircut', label: 'When was your last haircut?', type: 'select', options: ['This week', 'This month', 'Few months ago', 'I have stopped counting'] },
  { id: 'unread', label: 'How many unread messages right now?', type: 'select', options: ['Zero, I am a machine', 'Under 10', '10 to 100', 'Over 1000'] },
  { id: 'replyK', label: 'Do you reply with just "k"?', type: 'select', options: ['Never', 'Sometimes', 'Often', 'It is my whole personality'] },
  { id: 'mirror', label: 'Mirror selfies in your camera roll?', type: 'select', options: ['None', 'A few', 'Dozens', 'That is all there is'] },
  { id: 'opener', label: 'Your usual opening line', type: 'select', options: ['hey', 'hi :)', 'A real question about them', 'I wait for them to text first'] },
  { id: 'gym', label: 'Gym visits per week', type: 'select', options: ['0', '1 to 2', '3 to 4', '5 or more'] },
  { id: 'fit', label: 'Pick your daily outfit', type: 'select', options: ['Same hoodie every day', 'Plain shirt and jeans', 'Actually put together', 'Full drip, people stare'] },
  { id: 'lastDate', label: 'Last time you went on a date', type: 'select', options: ['This month', 'This year', 'Years ago', 'Define date'] },
  { id: 'confidence', label: 'Rate your own confidence', type: 'select', options: ['Low', 'Medium', 'High', 'Delusional'] }
]

const SYSTEM_PROMPT = `You are Rizz Score, a comedy judge for a joke web app made by university students.
A user answers ten silly questions about themselves. You rate their "rizz" (charisma) from 0 to 100.

Rules:
- Be funny and a bit brutal, like a friend roasting them. Never cruel about body, race, religion, or anything a person cannot change.
- The roast is one or two short sentences. Plain simple English.
- Give a short title that sounds like a rank, for example "Certified Yapper" or "Silent Mode Enjoyer".
- Give exactly three short tips. Two should be jokes, one should be quietly useful.
- Vary the score. Do not always land near 50.
- The user may also have done a face scan. You never see the face, only whether they pulled a funny face, sat still, or skipped it. Work that into the roast in one clause. A funny face is brave and earns a small bonus. Sitting frozen is funnier to mock. Skipping means they were scared.

Reply with JSON only, no markdown fence, in this exact shape:
{"score": 0, "title": "", "roast": "", "tips": ["", "", ""]}`

function fallbackScore(answers) {
  // Used only when the LLM call fails, so a live demo never shows a blank screen.
  let score = 50
  for (const q of QUESTIONS) {
    const i = q.options.indexOf(answers[q.id])
    if (i >= 0) score += (i - 1.5) * 4
  }
  score = Math.max(1, Math.min(99, Math.round(score)))
  return {
    score,
    title: 'Offline Rizz',
    roast: 'The judge is not answering right now, so you get the boring maths version.',
    tips: ['Try again in a moment', 'Check the API key', 'Touch grass while you wait'],
    offline: true
  }
}

const FACE_NOTES = {
  mocked: 'Face scan: they pulled a ridiculous face on camera.',
  still: 'Face scan: they sat completely still and stared at the lens.',
  skipped: 'Face scan: they refused to turn the camera on.'
}

async function askLLM(answers, face) {
  const lines = [
    ...QUESTIONS.map(q => `${q.label} ${answers[q.id] || 'no answer'}`),
    FACE_NOTES[face] || FACE_NOTES.skipped
  ].join('\n')

  const res = await fetch(`${BASE_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${API_KEY}`
    },
    body: JSON.stringify({
      model: MODEL,
      temperature: 1,
      reasoning_effort: 'low',
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: lines }
      ]
    })
  })

  if (!res.ok) throw new Error(`LLM responded ${res.status}: ${await res.text()}`)

  const data = await res.json()
  // Reasoning models sometimes wrap the JSON in a code fence, so take the
  // outermost braces rather than trusting the whole string.
  const raw = data.choices[0].message.content || ''
  const body = raw.slice(raw.indexOf('{'), raw.lastIndexOf('}') + 1)
  const parsed = JSON.parse(body)

  return {
    score: Math.max(0, Math.min(100, Number(parsed.score) || 0)),
    title: String(parsed.title || 'Unrated'),
    roast: String(parsed.roast || ''),
    tips: Array.isArray(parsed.tips) ? parsed.tips.slice(0, 3).map(String) : []
  }
}

app.get('/api/questions', (req, res) => res.json(QUESTIONS))

app.post('/api/rizz', async (req, res) => {
  const answers = req.body?.answers || {}
  const face = req.body?.face

  if (!API_KEY) return res.json(fallbackScore(answers))

  try {
    res.json(await askLLM(answers, face))
  } catch (err) {
    console.error('LLM call failed:', err.message)
    res.json(fallbackScore(answers))
  }
})

app.listen(PORT, () => console.log(`Rizz Score running on http://localhost:${PORT}`))

export default app
