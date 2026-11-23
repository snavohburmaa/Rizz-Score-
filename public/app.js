const form = document.getElementById('form')
const quiz = document.getElementById('quiz')
const scan = document.getElementById('scan')
const loading = document.getElementById('loading')
const result = document.getElementById('result')
const submit = document.getElementById('submit')
const video = document.getElementById('video')
const timerEl = document.getElementById('timer')
const promptEl = document.getElementById('prompt')
const meterFill = document.getElementById('meterfill')

const SCAN_MS = 7000
const MOCK_AT = 3000      // when we ask them to pull a face
const MOVE_LIMIT = 9      // average pixel change that counts as "they moved"

const answers = {}
let questions = []
let stream = null
let finished = false

async function init() {
  questions = await fetch('/api/questions').then(r => r.json())

  form.innerHTML = questions.map(q => `
    <div class="q">
      <label>${q.label}</label>
      <div class="opts" data-q="${q.id}">
        ${q.options.map(o => `<button type="button" class="opt" data-v="${o}">${o}</button>`).join('')}
      </div>
    </div>
  `).join('')

  form.addEventListener('click', e => {
    const btn = e.target.closest('.opt')
    if (!btn) return
    const group = btn.parentElement
    group.querySelectorAll('.opt').forEach(b => b.classList.remove('on'))
    btn.classList.add('on')
    answers[group.dataset.q] = btn.dataset.v
    submit.disabled = Object.keys(answers).length < questions.length
  })

  submit.disabled = true
}

function show(section) {
  for (const s of [quiz, scan, loading, result]) s.classList.add('hidden')
  section.classList.remove('hidden')
}

function stopCamera() {
  if (!stream) return
  stream.getTracks().forEach(t => t.stop())
  stream = null
}

// Compares downscaled frames and returns how much the picture changed.
// The face never leaves this function, and no frame is ever sent anywhere.
function makeMotionProbe() {
  const c = document.createElement('canvas')
  c.width = 64
  c.height = 48
  const ctx = c.getContext('2d', { willReadFrequently: true })
  let previous = null

  return () => {
    ctx.drawImage(video, 0, 0, c.width, c.height)
    const frame = ctx.getImageData(0, 0, c.width, c.height).data
    if (!previous) {
      previous = frame.slice()
      return 0
    }
    let total = 0
    for (let i = 0; i < frame.length; i += 4) total += Math.abs(frame[i] - previous[i])
    previous = frame.slice()
    return total / (frame.length / 4)
  }
}

async function startScan() {
  finished = false
  show(scan)
  scan.classList.remove('scanning')
  meterFill.style.width = '0'
  timerEl.textContent = Math.round(SCAN_MS / 1000)
  promptEl.textContent = 'Starting camera...'

  try {
    stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: 'user', width: { ideal: 640 } },
      audio: false
    })
  } catch (err) {
    promptEl.textContent = 'No camera, no problem. Judging you on words alone.'
    setTimeout(() => finishScan('skipped'), 1200)
    return
  }

  video.srcObject = stream
  scan.classList.add('scanning')

  const probe = makeMotionProbe()
  const started = performance.now()
  let peak = 0
  let asked = false

  promptEl.textContent = 'Hold still. Reading your aura through the lens.'

  const tick = setInterval(() => {
    if (finished) return clearInterval(tick)

    const elapsed = performance.now() - started
    const left = Math.max(0, SCAN_MS - elapsed)

    timerEl.textContent = Math.ceil(left / 1000)
    meterFill.style.width = Math.min(100, (elapsed / SCAN_MS) * 100) + '%'

    if (!asked && elapsed > MOCK_AT) {
      asked = true
      promptEl.textContent = 'Now pull your silliest face. Or do not. We are watching either way.'
    }

    // Only the window after the prompt counts, so walking up to the laptop
    // does not get scored as a funny face.
    if (asked) peak = Math.max(peak, probe())

    if (left <= 0) {
      clearInterval(tick)
      promptEl.textContent = 'Scan complete.'
      finishScan(peak > MOVE_LIMIT ? 'mocked' : 'still')
    }
  }, 100)
}

async function finishScan(face) {
  if (finished) return
  finished = true
  scan.classList.remove('scanning')
  stopCamera()
  await sendScore(face)
}

async function sendScore(face) {
  show(loading)

  const res = await fetch('/api/rizz', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ answers, face })
  }).then(r => r.json())

  document.getElementById('rank').textContent = res.title
  document.getElementById('roast').textContent = res.roast
  document.getElementById('tips').innerHTML = res.tips.map(t => `<li>${t}</li>`).join('')

  show(result)

  // Count the number up so the reveal feels like a game show.
  const el = document.getElementById('score')
  let n = 0
  const step = Math.max(1, Math.round(res.score / 40))
  const timer = setInterval(() => {
    n = Math.min(res.score, n + step)
    el.textContent = n
    if (n >= res.score) clearInterval(timer)
  }, 25)

  requestAnimationFrame(() => {
    document.getElementById('fill').style.width = res.score + '%'
  })
}

submit.addEventListener('click', startScan)

document.getElementById('skip').addEventListener('click', () => finishScan('skipped'))

document.getElementById('again').addEventListener('click', () => {
  document.getElementById('fill').style.width = '0'
  show(quiz)
})

init()
