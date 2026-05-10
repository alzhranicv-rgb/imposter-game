let wordsList = []
let selectedItem = null

let players = []
let imposterIndex = null
let currentViewedIndex = null

let turnDrawRunning = false
let turnDrawTimer = null

let usedWordKeys = []

const FIXED_PLAYERS_STORAGE_KEY = "imposter_fixed_players_v3"
const USED_WORDS_STORAGE_KEY = "imposter_used_words_v1"
const GAME_SETTINGS_STORAGE_KEY = "imposter_game_settings_v1"

/* =========================
   تحميل ملف Excel
========================= */

async function loadExcelFromProject() {
  const fileInfo = document.getElementById("fileInfo")

  try {
    const response = await fetch("categorized_words.xlsx")

    if (!response.ok) {
      throw new Error("لم يتم العثور على ملف Excel")
    }

    const arrayBuffer = await response.arrayBuffer()
    const data = new Uint8Array(arrayBuffer)

    const workbook = XLSX.read(data, {
      type: "array"
    })

    let allItems = []

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName]

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
      })

      const sheetItems = extractWordsFromRows(rows)
      allItems = allItems.concat(sheetItems)
    })

    wordsList = allItems
    loadUsedWords()

    if (wordsList.length === 0) {
      fileInfo.textContent = "لم يتم تحميل كلمات"
      showWarning("ملف Excel لا يحتوي على كلمات واضحة")
      return
    }

    fileInfo.textContent = `تم تحميل ${wordsList.length} كلمة تلقائيًا`
    hideWarning()

  } catch (error) {
    fileInfo.textContent = "تعذر تحميل ملف الكلمات"
    showWarning("تأكد أن ملف categorized_words.xlsx موجود بجانب index.html")
    console.error(error)
  }
}

function extractWordsFromRows(rows) {
  const result = []
  let lastCategory = ""

  rows.forEach((row, rowIndex) => {
    if (rowIndex === 0) return

    let category = String(row[0] || "").trim()
    const word = String(row[1] || "").trim()
    const description = String(row[2] || "").trim()

    if (isHeaderText(category) || isHeaderText(word)) return

    if (category) {
      lastCategory = category
    } else {
      category = lastCategory
    }

    if (!word) return

    result.push({
      category: category || "بدون فئة",
      word,
      description
    })
  })

  return result
}

function isHeaderText(value) {
  const text = String(value || "").trim().toLowerCase()

  return (
    text === "الفئة" ||
    text === "التصنيف" ||
    text === "category" ||
    text === "الكلمة" ||
    text === "كلمات" ||
    text === "word" ||
    text === "words" ||
    text === "الوصف" ||
    text === "description"
  )
}

/* =========================
   اختيار الكلمات بدون تكرار
========================= */

function getWordKey(item) {
  return `${item.category}__${item.word}`
}

function loadUsedWords() {
  try {
    const saved = JSON.parse(localStorage.getItem(USED_WORDS_STORAGE_KEY) || "[]")

    if (Array.isArray(saved)) {
      usedWordKeys = saved
    }
  } catch {
    usedWordKeys = []
  }
}

function saveUsedWords() {
  localStorage.setItem(USED_WORDS_STORAGE_KEY, JSON.stringify(usedWordKeys))
}

function getRandomItem() {
  if (wordsList.length === 0) return null

  let availableWords = wordsList.filter((item) => {
    return !usedWordKeys.includes(getWordKey(item))
  })

  if (availableWords.length === 0) {
    usedWordKeys = []
    saveUsedWords()
    availableWords = [...wordsList]
  }

  const randomIndex = Math.floor(Math.random() * availableWords.length)
  const selected = availableWords[randomIndex]

  usedWordKeys.push(getWordKey(selected))
  saveUsedWords()

  return selected
}

/* =========================
   اللاعبين الثابتين
========================= */

function getDefaultFixedPlayers() {
  return [
    "امبارك",
    "مشاري",
    "ابو عزة",
    "محمد سالم",
    "ابو ميلا",
    "نايف",
    "محمد عبدالله"
  ]
}

function getFixedPlayers() {
  try {
    const saved = JSON.parse(localStorage.getItem(FIXED_PLAYERS_STORAGE_KEY) || "null")

    if (Array.isArray(saved) && saved.length > 0) {
      return saved
    }
  } catch (error) {
    console.warn("تعذر قراءة قائمة اللاعبين", error)
  }

  return getDefaultFixedPlayers()
}

function saveFixedPlayers(names) {
  localStorage.setItem(FIXED_PLAYERS_STORAGE_KEY, JSON.stringify(names))
}

function renderFixedPlayers() {
  const list = document.getElementById("fixedPlayersList")
  const badge = document.getElementById("playersCountBadge")

  const names = getFixedPlayers()

  list.innerHTML = ""

  names.forEach((name, index) => {
    const item = document.createElement("div")
    item.className = "fixedPlayerItem"

    const nameEl = document.createElement("div")
    nameEl.className = "fixedPlayerName"
    nameEl.textContent = name

    const deleteBtn = document.createElement("button")
    deleteBtn.type = "button"
    deleteBtn.className = "deleteFixedPlayerBtn"
    deleteBtn.textContent = "×"
    deleteBtn.onclick = () => deleteFixedPlayer(index)

    item.appendChild(nameEl)
    item.appendChild(deleteBtn)

    list.appendChild(item)
  })

  badge.textContent = `${names.length} لاعبين`
}

function addFixedPlayer() {
  hideWarning()

  const input = document.getElementById("newPlayerName")
  const newName = input.value.trim()

  if (!newName) {
    showWarning("اكتب اسم اللاعب أولًا")
    return
  }

  const names = getFixedPlayers()

  if (names.length >= 20) {
    showWarning("الحد الأعلى 20 لاعب")
    return
  }

  const alreadyExists = names.some((name) => {
    return name.trim() === newName
  })

  if (alreadyExists) {
    showWarning("هذا الاسم موجود مسبقًا")
    return
  }

  names.push(newName)
  saveFixedPlayers(names)

  input.value = ""
  renderFixedPlayers()
}

function deleteFixedPlayer(index) {
  hideWarning()

  const names = getFixedPlayers()

  if (names.length <= 3) {
    showWarning("لازم يبقى عندك 3 لاعبين على الأقل")
    return
  }

  names.splice(index, 1)
  saveFixedPlayers(names)
  renderFixedPlayers()
}

/* =========================
   إعدادات اللعبة
========================= */

function getDefaultGameSettings() {
  return {
    showCategoryForImposter: true
  }
}

function getGameSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(GAME_SETTINGS_STORAGE_KEY) || "null")

    if (saved && typeof saved === "object") {
      return {
        ...getDefaultGameSettings(),
        ...saved
      }
    }
  } catch (error) {
    console.warn("تعذر قراءة إعدادات اللعبة", error)
  }

  return getDefaultGameSettings()
}

function saveGameSettings() {
  const checkbox = document.getElementById("showCategoryForImposter")

  const settings = {
    showCategoryForImposter: checkbox ? checkbox.checked : true
  }

  localStorage.setItem(GAME_SETTINGS_STORAGE_KEY, JSON.stringify(settings))
}

function applyGameSettingsToUI() {
  const settings = getGameSettings()
  const checkbox = document.getElementById("showCategoryForImposter")

  if (checkbox) {
    checkbox.checked = settings.showCategoryForImposter
  }
}

/* =========================
   بداية اللعبة
========================= */

function startGame() {
  const names = getFixedPlayers()

  if (wordsList.length === 0) {
    showWarning("ملف الكلمات لم يتم تحميله بعد")
    return
  }

  if (names.length < 3) {
    showWarning("لازم يكون عدد اللاعبين 3 أو أكثر")
    return
  }

  selectedItem = getRandomItem()

  if (!selectedItem) {
    showWarning("تعذر اختيار كلمة")
    return
  }

  players = []

  names.forEach((name) => {
    players.push({
      name,
      category: selectedItem.category,
      word: selectedItem.word,
      description: selectedItem.description,
      isImposter: false,
      viewed: false
    })
  })

  imposterIndex = Math.floor(Math.random() * players.length)

  players[imposterIndex].isImposter = true
  players[imposterIndex].word = ""

  renderCards()
  updateCounter()
  resetTurnBox()
  updateGameButtons()

  showScreen("gameScreen")
}

function renderCards() {
  const cardsGrid = document.getElementById("cardsGrid")
  cardsGrid.innerHTML = ""

  players.forEach((player, index) => {
    const card = document.createElement("button")

    card.className = `playerCard ${player.viewed ? "viewed" : ""}`
    card.textContent = player.viewed ? `${player.name} ✓` : player.name
    card.onclick = () => showSecret(index)

    cardsGrid.appendChild(card)
  })
}

function showSecret(index) {
  currentViewedIndex = index

  const player = players[index]
  const settings = getGameSettings()

  const secretName = document.getElementById("secretName")
  const secretCategory = document.getElementById("secretCategory")
  const secretWord = document.getElementById("secretWord")

  secretName.textContent = player.name

  if (player.isImposter) {
    secretWord.textContent = "أنت الإمبوستر"
    secretWord.classList.add("imposterWord")

    if (settings.showCategoryForImposter) {
      secretCategory.textContent = `الفئة: ${player.category}`
      secretCategory.style.display = "inline-flex"
    } else {
      secretCategory.textContent = ""
      secretCategory.style.display = "none"
    }
  } else {
    secretCategory.textContent = `الفئة: ${player.category}`
    secretCategory.style.display = "inline-flex"

    secretWord.textContent = player.word
    secretWord.classList.remove("imposterWord")
  }

  document.getElementById("secretScreen").style.display = "flex"
}

function hideSecret() {
  if (currentViewedIndex !== null) {
    players[currentViewedIndex].viewed = true
  }

  currentViewedIndex = null

  document.getElementById("secretScreen").style.display = "none"

  renderCards()
  updateCounter()
  updateGameButtons()
}

/* =========================
   قرعة الدور
========================= */

function startTurnDraw() {
  if (!players.length || turnDrawRunning) return

  const screen = document.getElementById("turnDrawScreen")
  const nameBox = document.getElementById("turnDrawName")
  const hint = document.getElementById("turnDrawHint")
  const closeBtn = document.getElementById("closeTurnDrawBtn")

  if (!screen || !nameBox || !hint || !closeBtn) return

  turnDrawRunning = true

  screen.style.display = "flex"
  nameBox.classList.remove("winner")
  closeBtn.classList.add("hidden")
  hint.textContent = "جاري اختيار اللاعب..."

  let index = 0
  let loops = 0

  if (turnDrawTimer) {
    clearInterval(turnDrawTimer)
  }

  turnDrawTimer = setInterval(() => {
    nameBox.textContent = players[index].name

    index++

    if (index >= players.length) {
      index = 0
      loops++
    }

    if (loops >= 5) {
      clearInterval(turnDrawTimer)
      turnDrawTimer = null

      const winnerIndex = Math.floor(Math.random() * players.length)
      const winner = players[winnerIndex]

      setTimeout(() => {
        nameBox.textContent = winner.name
        nameBox.classList.add("winner")
        hint.textContent = "الدور عليه"
        closeBtn.classList.remove("hidden")
        turnDrawRunning = false
      }, 280)
    }
  }, 65)
}

function closeTurnDrawScreen() {
  if (turnDrawRunning) return

  const screen = document.getElementById("turnDrawScreen")

  if (screen) {
    screen.style.display = "none"
  }
}

function resetTurnBox() {
  const screen = document.getElementById("turnDrawScreen")
  const nameBox = document.getElementById("turnDrawName")
  const hint = document.getElementById("turnDrawHint")
  const closeBtn = document.getElementById("closeTurnDrawBtn")

  if (turnDrawTimer) {
    clearInterval(turnDrawTimer)
    turnDrawTimer = null
  }

  turnDrawRunning = false

  if (screen) screen.style.display = "none"

  if (nameBox) {
    nameBox.textContent = "جاهز؟"
    nameBox.classList.remove("winner")
  }

  if (hint) {
    hint.textContent = "انتظر حتى تتوقف القرعة"
  }

  if (closeBtn) {
    closeBtn.classList.add("hidden")
  }
}

/* =========================
   أزرار نهاية الجولة
========================= */

function updateGameButtons() {
  const revealBtn = document.getElementById("revealBtn")
  const newRoundBtn = document.getElementById("newRoundBtn")
  const resetBtn = document.getElementById("resetBtn")

  if (!revealBtn || !newRoundBtn || !resetBtn) return

  const allViewed = players.length > 0 && players.every((player) => player.viewed)

  if (allViewed) {
    revealBtn.classList.remove("hidden")
  } else {
    revealBtn.classList.add("hidden")
  }

  newRoundBtn.classList.add("hidden")
  resetBtn.classList.add("hidden")
}

function revealImposter() {
  if (!players.length || imposterIndex === null) return

  const imposter = players[imposterIndex]

  document.getElementById("revealName").textContent = imposter.name

  document.getElementById("revealDetails").innerHTML = `
    الفئة: ${selectedItem.category}
    <br>
    الكلمة: ${selectedItem.word}
  `

  document.getElementById("revealScreen").style.display = "flex"

  const revealBtn = document.getElementById("revealBtn")
  const newRoundBtn = document.getElementById("newRoundBtn")
  const resetBtn = document.getElementById("resetBtn")

  if (revealBtn) revealBtn.classList.add("hidden")
  if (newRoundBtn) newRoundBtn.classList.remove("hidden")
  if (resetBtn) resetBtn.classList.remove("hidden")
}

function closeReveal() {
  document.getElementById("revealScreen").style.display = "none"
}

function newRound() {
  if (players.length === 0) return

  selectedItem = getRandomItem()

  if (!selectedItem) {
    showWarning("تعذر اختيار كلمة")
    return
  }

  players = players.map((player) => {
    return {
      name: player.name,
      category: selectedItem.category,
      word: selectedItem.word,
      description: selectedItem.description,
      isImposter: false,
      viewed: false
    }
  })

  imposterIndex = Math.floor(Math.random() * players.length)

  players[imposterIndex].isImposter = true
  players[imposterIndex].word = ""

  renderCards()
  updateCounter()
  closeReveal()
  resetTurnBox()
  updateGameButtons()
}

function resetGame() {
  players = []
  imposterIndex = null
  currentViewedIndex = null
  selectedItem = null

  document.getElementById("secretScreen").style.display = "none"
  document.getElementById("revealScreen").style.display = "none"

  resetTurnBox()
  updateGameButtons()
  renderFixedPlayers()
  showScreen("setupScreen")
}

/* =========================
   أدوات عامة
========================= */

function updateCounter() {
  const viewed = players.filter((player) => player.viewed).length
  const total = players.length

  document.getElementById("viewedCounter").textContent = viewed
  document.getElementById("totalCounter").textContent = total
}

function showScreen(screenId) {
  document.querySelectorAll(".screen").forEach((screen) => {
    screen.classList.remove("active")
  })

  document.getElementById(screenId).classList.add("active")
}

function showWarning(text) {
  const warning = document.getElementById("warning")

  warning.textContent = text
  warning.style.display = "block"
}

function hideWarning() {
  const warning = document.getElementById("warning")

  warning.textContent = ""
  warning.style.display = "none"
}

/* =========================
   تشغيل أولي
========================= */

renderFixedPlayers()
applyGameSettingsToUI()
loadExcelFromProject()