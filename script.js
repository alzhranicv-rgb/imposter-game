let wordsList = []
let selectedItem = null

let players = []
let imposterIndex = null
let currentViewedIndex = null

let turnDrawRunning = false
let turnDrawTimer = null
let turnDrawDone = false
let votingStarted = false
let turnDrawPool = []
let lastTurnWinnerName = ""

let usedWordKeys = []

let votes = {}
let scores = {}
let imposterRevealed = false
let roundScored = false
let currentVoteIndex = 0

const FIXED_PLAYERS_STORAGE_KEY = "imposter_fixed_players_v3"
const USED_WORDS_STORAGE_KEY = "imposter_used_words_v1"
const GAME_SETTINGS_STORAGE_KEY = "imposter_game_settings_v1"
const GAME_STATE_STORAGE_KEY = "imposter_live_game_state_v1"

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
      if (fileInfo) fileInfo.textContent = "لم يتم تحميل كلمات"
      showWarning("ملف Excel لا يحتوي على كلمات واضحة")
      return
    }

    if (fileInfo) fileInfo.textContent = `تم تحميل ${wordsList.length} كلمة تلقائيًا`
    hideWarning()

    restoreGameState()
  } catch (error) {
    if (fileInfo) fileInfo.textContent = "تعذر تحميل ملف الكلمات"
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
   حفظ واستعادة حالة اللعبة
========================= */

function saveGameState() {
  if (!players || players.length === 0 || !selectedItem) return

  const state = {
    selectedItem,
    players,
    imposterIndex,
    votes,
    scores,
    imposterRevealed,
    roundScored,
    currentVoteIndex,
    turnDrawDone,
    votingStarted,
    turnDrawPool,
    lastTurnWinnerName
  }

  localStorage.setItem(GAME_STATE_STORAGE_KEY, JSON.stringify(state))
}

function clearGameState() {
  localStorage.removeItem(GAME_STATE_STORAGE_KEY)
}

function restoreGameState() {
  try {
    const saved = JSON.parse(localStorage.getItem(GAME_STATE_STORAGE_KEY) || "null")

    if (!saved || !Array.isArray(saved.players) || saved.players.length === 0) {
      renderFixedPlayers()
      applyGameSettingsToUI()
      showScreen("setupScreen")
      return
    }

    selectedItem = saved.selectedItem || null
    players = saved.players || []
    imposterIndex = typeof saved.imposterIndex === "number" ? saved.imposterIndex : null
    votes = saved.votes || {}
    scores = saved.scores || {}
    imposterRevealed = !!saved.imposterRevealed
    roundScored = !!saved.roundScored
    currentVoteIndex = Number(saved.currentVoteIndex || 0)
    turnDrawDone = !!saved.turnDrawDone
    votingStarted = !!saved.votingStarted
    turnDrawPool = Array.isArray(saved.turnDrawPool) ? saved.turnDrawPool : []
    lastTurnWinnerName = saved.lastTurnWinnerName || ""

    renderFixedPlayers()
    applyGameSettingsToUI()
    showScreen("gameScreen")
    renderRestoredGame()
  } catch (error) {
    console.warn("تعذر استعادة حالة اللعبة", error)
    clearGameState()
    showScreen("setupScreen")
  }
}

function renderRestoredGame() {
  closeAllOverlaysInstant()

  renderCards()
  updateCounter()
  renderScoreBoard()
  updateGameStageAfterRestore()
  updateGameButtons()
}

function updateGameStageAfterRestore() {
  const gameStageTitle = document.getElementById("gameStageTitle")
  const turnBox = document.getElementById("turnBox")
  const drawTurnBtn = document.getElementById("drawTurnBtn")
  const startVotingBtn = document.getElementById("startVotingBtn")
  const votingPanel = document.getElementById("votingPanel")
  const settings = getGameSettings()

  const turnBoxIcon = document.querySelector(".turnBoxIcon")
  const turnBoxTitle = document.querySelector(".turnBoxText h3")
  const turnBoxText = document.querySelector(".turnBoxText p")

  if (turnBox) {
    turnBox.classList.add("hidden")
    turnBox.classList.remove("drawStage", "voteReadyStage", "stageLeaving")
  }

  if (drawTurnBtn) drawTurnBtn.classList.add("hidden")
  if (startVotingBtn) startVotingBtn.classList.add("hidden")
  if (votingPanel) votingPanel.classList.add("hidden")

  if (!areAllCardsViewed()) {
    if (gameStageTitle) gameStageTitle.textContent = "اختار اسمك وشوف بطاقتك بسرية"
    return
  }

  if (!turnDrawDone && !imposterRevealed) {
    if (turnBox) {
      turnBox.classList.remove("hidden")
      turnBox.classList.add("drawStage")
    }

    if (turnBoxIcon) turnBoxIcon.textContent = "🎲"
    if (turnBoxTitle) turnBoxTitle.textContent = "مرحلة القرعة"
    if (turnBoxText) {
      turnBoxText.textContent = "كل البطاقات انفتحت. الآن اختار اللاعب اللي يبدأ الكلام بدون ما ينكشف السر."
    }

    if (drawTurnBtn) {
      drawTurnBtn.classList.remove("hidden")
      drawTurnBtn.textContent = "ابدأ القرعة"
    }

    if (gameStageTitle) gameStageTitle.textContent = "كل البطاقات انفتحت"
    return
  }

  if (turnDrawDone && !votingStarted && !imposterRevealed) {
    if (turnBox) {
      turnBox.classList.remove("hidden")
      turnBox.classList.add("voteReadyStage")
    }

    if (turnBoxIcon) turnBoxIcon.textContent = "🗳️"
    if (turnBoxTitle) turnBoxTitle.textContent = "جاهزين للتصويت؟"
    if (turnBoxText) {
      turnBoxText.textContent = lastTurnWinnerName
        ? `بدأ الدور على ${lastTurnWinnerName}. بعد النقاش، ابدأوا التصويت لاختيار الإمبوستر.`
        : "بعد النقاش، ابدأوا التصويت لاختيار الإمبوستر."
    }

    if (startVotingBtn) {
      startVotingBtn.classList.remove("hidden")
      startVotingBtn.textContent = "بدء التصويت"
    }

    if (gameStageTitle) gameStageTitle.textContent = "مرحلة ما قبل التصويت"
    return
  }

  if (imposterRevealed) {
    if (gameStageTitle) gameStageTitle.textContent = "تم كشف الإمبوستر"
    return
  }

  if (settings.enableVoting && votingStarted && !isVotingComplete()) {
    if (gameStageTitle) gameStageTitle.textContent = "مرحلة التصويت"
    renderCurrentVote()
    return
  }

  if (gameStageTitle) gameStageTitle.textContent = "جاهز لكشف الإمبوستر"
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

  if (!list || !badge) return

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
  if (!input) return

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
    showCategoryForImposter: true,
    enableVoting: true
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
  const showCategoryCheckbox = document.getElementById("showCategoryForImposter")
  const enableVotingCheckbox = document.getElementById("enableVoting")

  const settings = {
    showCategoryForImposter: showCategoryCheckbox ? showCategoryCheckbox.checked : true,
    enableVoting: enableVotingCheckbox ? enableVotingCheckbox.checked : true
  }

  localStorage.setItem(GAME_SETTINGS_STORAGE_KEY, JSON.stringify(settings))

  if (players.length > 0) {
    updateGameStageAfterRestore()
    updateGameButtons()
    saveGameState()
  }
}

function applyGameSettingsToUI() {
  const settings = getGameSettings()

  const showCategoryCheckbox = document.getElementById("showCategoryForImposter")
  const enableVotingCheckbox = document.getElementById("enableVoting")

  if (showCategoryCheckbox) {
    showCategoryCheckbox.checked = settings.showCategoryForImposter
  }

  if (enableVotingCheckbox) {
    enableVotingCheckbox.checked = settings.enableVoting
  }
}

/* =========================
   التصويت والنقاط
========================= */

function initializeScores(names) {
  scores = {}

  names.forEach((name) => {
    scores[name] = 0
  })

  renderScoreBoard()
}

function resetRoundVoting() {
  votes = {}
  currentVoteIndex = 0
  imposterRevealed = false
  roundScored = false
  turnDrawDone = false
  votingStarted = false
  lastTurnWinnerName = ""

  const votingPanel = document.getElementById("votingPanel")
  const votingChoices = document.getElementById("votingChoices")
  const nextVoteBtn = document.getElementById("nextVoteBtn")
  const startVotingBtn = document.getElementById("startVotingBtn")

  if (votingPanel) votingPanel.classList.add("hidden")
  if (votingChoices) votingChoices.innerHTML = ""
  if (nextVoteBtn) nextVoteBtn.classList.add("hidden")
  if (startVotingBtn) startVotingBtn.classList.add("hidden")
}

function renderCurrentVote() {
  const settings = getGameSettings()
  const votingPanel = document.getElementById("votingPanel")
  const votingChoices = document.getElementById("votingChoices")
  const currentVoterTitle = document.getElementById("currentVoterTitle")
  const nextVoteBtn = document.getElementById("nextVoteBtn")

  if (!votingPanel || !votingChoices || !currentVoterTitle || !nextVoteBtn) return

  if (!settings.enableVoting || !turnDrawDone || imposterRevealed) {
    votingPanel.classList.add("hidden")
    nextVoteBtn.classList.add("hidden")
    return
  }

  if (currentVoteIndex >= players.length) {
    votingPanel.classList.add("hidden")
    nextVoteBtn.classList.add("hidden")
    updateGameButtons()
    return
  }

  const voter = players[currentVoteIndex]
  const isLastVoter = currentVoteIndex === players.length - 1

  votingPanel.classList.remove("hidden")
  currentVoterTitle.textContent = `تصويت ${voter.name}`
  votingChoices.innerHTML = ""
  nextVoteBtn.classList.add("hidden")

  players.forEach((target) => {
    const btn = document.createElement("button")
    btn.type = "button"
    btn.className = "voteChoiceBtn"
    btn.textContent = target.name

    if (votes[voter.name] === target.name) {
      btn.classList.add("selected")

      if (!isLastVoter) {
        nextVoteBtn.classList.remove("hidden")
      }
    }

    btn.onclick = () => {
      votes[voter.name] = target.name

      document.querySelectorAll(".voteChoiceBtn").forEach((choiceBtn) => {
        choiceBtn.classList.remove("selected")
      })

      btn.classList.add("selected")

      if (isLastVoter) {
        votingPanel.classList.add("hidden")
        nextVoteBtn.classList.add("hidden")

        const gameStageTitle = document.getElementById("gameStageTitle")
        if (gameStageTitle) gameStageTitle.textContent = "جاهز لكشف الإمبوستر"

        updateGameButtons()
        saveGameState()
        return
      }

      nextVoteBtn.classList.remove("hidden")

      saveGameState()
      updateGameButtons()
    }

    votingChoices.appendChild(btn)
  })

  updateGameButtons()
}

function goToNextVoter() {
  if (currentVoteIndex >= players.length) return

  const voter = players[currentVoteIndex]

  if (!votes[voter.name]) {
    showWarning("اختر اسم قبل الانتقال للي بعده")
    return
  }

  hideWarning()

  currentVoteIndex++

  if (currentVoteIndex >= players.length) {
    const votingPanel = document.getElementById("votingPanel")
    const nextVoteBtn = document.getElementById("nextVoteBtn")
    const gameStageTitle = document.getElementById("gameStageTitle")

    if (votingPanel) votingPanel.classList.add("hidden")
    if (nextVoteBtn) nextVoteBtn.classList.add("hidden")
    if (gameStageTitle) gameStageTitle.textContent = "جاهز لكشف الإمبوستر"

    updateGameButtons()
    saveGameState()
    return
  }

  renderCurrentVote()
  updateGameButtons()
  saveGameState()
}

function isVotingComplete() {
  const settings = getGameSettings()

  if (!settings.enableVoting) return true

  return players.length > 0 && players.every((player) => {
    return !!votes[player.name]
  })
}

function calculateRoundScores() {
  if (roundScored) return ""

  const settings = getGameSettings()
  const imposterName = players[imposterIndex].name
  const correctVoters = []

  if (settings.enableVoting) {
    Object.keys(votes).forEach((voterName) => {
      if (voterName === imposterName) return

      if (votes[voterName] === imposterName) {
        correctVoters.push(voterName)
      }
    })
  }

  let resultText = ""

  if (!settings.enableVoting) {
    resultText = "التصويت غير مفعّل في هذه الجولة"
  } else if (correctVoters.length > 0) {
    correctVoters.forEach((name) => {
      scores[name] = (scores[name] || 0) + 1
    })

    resultText = `اللي صوتوا صح: ${correctVoters.join("، ")}`
  } else {
    scores[imposterName] = (scores[imposterName] || 0) + 2
    resultText = `ما أحد صوت على الإمبوستر، ${imposterName} أخذ نقطتين`
  }

  roundScored = true
  renderScoreBoard()
  saveGameState()

  return resultText
}

function renderScoreBoard() {
  const scoreList = document.getElementById("scoreList")

  if (!scoreList) return

  const names = Object.keys(scores)

  scoreList.innerHTML = ""

  if (names.length === 0) return

  const sortedNames = names.sort((a, b) => {
    return (scores[b] || 0) - (scores[a] || 0)
  })

  sortedNames.forEach((name, index) => {
    const item = document.createElement("div")
    item.className = "scoreItem"

    const nameEl = document.createElement("div")
    nameEl.className = "scoreName"
    nameEl.textContent = index === 0 ? `👑 ${name}` : name

    const pointsEl = document.createElement("div")
    pointsEl.className = "scorePoints"
    pointsEl.textContent = `${scores[name] || 0}`

    item.appendChild(nameEl)
    item.appendChild(pointsEl)

    scoreList.appendChild(item)
  })
}

function openScoreBoard() {
  renderScoreBoard()

  const scoreScreen = document.getElementById("scoreScreen")

  if (scoreScreen) {
    scoreScreen.classList.remove("closing")
    scoreScreen.classList.add("show")
    scoreScreen.style.display = "flex"
  }
}

function closeScoreBoard() {
  const scoreScreen = document.getElementById("scoreScreen")

  if (!scoreScreen) return

  scoreScreen.classList.remove("show")
  scoreScreen.classList.add("closing")

  setTimeout(() => {
    scoreScreen.classList.remove("closing")
    scoreScreen.style.display = "none"
  }, 180)
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

  initializeScores(names)
  turnDrawPool = shuffleNames(names)
  startNewRoundWithCurrentPlayers(names)

  showScreen("gameScreen")
  saveGameState()
}

function getRandomIndex(length) {
  if (length <= 0) return 0

  const randomArray = new Uint32Array(1)
  crypto.getRandomValues(randomArray)

  return randomArray[0] % length
}

function startNewRoundWithCurrentPlayers(names) {
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

  imposterIndex = getRandomIndex(players.length)

  players[imposterIndex].isImposter = true
  players[imposterIndex].word = ""

  resetRoundVoting()
  resetTurnBox()

  const cardsGrid = document.getElementById("cardsGrid")
  const turnBox = document.getElementById("turnBox")
  const gameStageTitle = document.getElementById("gameStageTitle")
  const playersCounter = document.querySelector(".playersCounter")

  if (cardsGrid) {
    cardsGrid.classList.remove("hidden", "stageLeaving")
  }

  if (turnBox) {
    turnBox.classList.add("hidden")
    turnBox.classList.remove("drawStage", "voteReadyStage", "stageLeaving")
  }

  if (gameStageTitle) gameStageTitle.textContent = "اختار اسمك وشوف بطاقتك بسرية"
  if (playersCounter) playersCounter.classList.remove("hidden")

  renderCards()
  updateCounter()
  renderScoreBoard()
  updateGameButtons()
  saveGameState()
}

function areAllCardsViewed() {
  return players.length > 0 && players.every((player) => player.viewed)
}

function renderCards() {
  const cardsGrid = document.getElementById("cardsGrid")
  const turnBox = document.getElementById("turnBox")
  const drawTurnBtn = document.getElementById("drawTurnBtn")
  const startVotingBtn = document.getElementById("startVotingBtn")
  const gameStageTitle = document.getElementById("gameStageTitle")
  const playersCounter = document.querySelector(".playersCounter")

  const turnBoxIcon = document.querySelector(".turnBoxIcon")
  const turnBoxTitle = document.querySelector(".turnBoxText h3")
  const turnBoxText = document.querySelector(".turnBoxText p")

  if (!cardsGrid || !turnBox) return

  const allViewed = areAllCardsViewed()

  cardsGrid.classList.remove(
    "playersCountSmall",
    "playersCountMedium",
    "playersCountLarge"
  )

  if (players.length <= 6) {
    cardsGrid.classList.add("playersCountSmall")
  } else if (players.length <= 8) {
    cardsGrid.classList.add("playersCountMedium")
  } else {
    cardsGrid.classList.add("playersCountLarge")
  }

  turnBox.classList.remove("drawStage", "voteReadyStage", "stageLeaving")

  if (allViewed) {
    cardsGrid.classList.add("stageLeaving")

    setTimeout(() => {
      cardsGrid.classList.add("hidden")
      cardsGrid.classList.remove("stageLeaving")
    }, 260)

    if (playersCounter) {
      playersCounter.classList.add("hidden")
    }

    if (!turnDrawDone && !imposterRevealed) {
      turnBox.classList.remove("hidden")
      turnBox.classList.add("drawStage")

      if (turnBoxIcon) turnBoxIcon.textContent = "🎲"
      if (turnBoxTitle) turnBoxTitle.textContent = "مرحلة القرعة"
      if (turnBoxText) {
        turnBoxText.textContent = "كل البطاقات انفتحت. الآن اختار اللاعب اللي يبدأ الكلام بدون ما ينكشف السر."
      }

      if (drawTurnBtn) {
        drawTurnBtn.classList.remove("hidden")
        drawTurnBtn.textContent = "ابدأ القرعة"
      }

      if (startVotingBtn) {
        startVotingBtn.classList.add("hidden")
      }

      if (gameStageTitle) {
        gameStageTitle.textContent = "كل البطاقات انفتحت"
      }
    }

    if (turnDrawDone && !votingStarted && !imposterRevealed) {
      turnBox.classList.remove("hidden")
      turnBox.classList.add("voteReadyStage")

      if (turnBoxIcon) turnBoxIcon.textContent = "🗳️"
      if (turnBoxTitle) turnBoxTitle.textContent = "جاهزين للتصويت؟"
      if (turnBoxText) {
        turnBoxText.textContent = lastTurnWinnerName
          ? `بدأ الدور على ${lastTurnWinnerName}. بعد النقاش، ابدأوا التصويت لاختيار الإمبوستر.`
          : "بعد النقاش، ابدأوا التصويت لاختيار الإمبوستر."
      }

      if (drawTurnBtn) {
        drawTurnBtn.classList.add("hidden")
      }

      if (startVotingBtn) {
        startVotingBtn.classList.remove("hidden")
        startVotingBtn.textContent = "بدء التصويت"
      }

      if (gameStageTitle) {
        gameStageTitle.textContent = "مرحلة ما قبل التصويت"
      }
    }

    return
  }

  cardsGrid.classList.remove("hidden", "stageLeaving")
  turnBox.classList.add("hidden")

  if (drawTurnBtn) {
    drawTurnBtn.classList.add("hidden")
  }

  if (startVotingBtn) {
    startVotingBtn.classList.add("hidden")
  }

  if (playersCounter) {
    playersCounter.classList.remove("hidden")
  }

  if (gameStageTitle) {
    gameStageTitle.textContent = "اختار اسمك وشوف بطاقتك بسرية"
  }

  cardsGrid.innerHTML = ""

  players.forEach((player, index) => {
    const card = document.createElement("button")

    card.type = "button"
    card.className = `playerCard ${player.viewed ? "viewed" : "locked"}`
    card.innerHTML = `
  <span>${player.name}</span>
`

    if (player.viewed) {
      card.disabled = true
    } else {
      card.onclick = () => openPlayerCard(index, card)
    }

    cardsGrid.appendChild(card)
  })
}

function openPlayerCard(index, card) {
  if (!card || card.classList.contains("opening")) return

  card.classList.add("opening")

  setTimeout(() => {
    showSecret(index)
    card.classList.remove("opening")
  }, 160)
}

function showSecret(index) {
  const player = players[index]

  if (!player || player.viewed) return

  currentViewedIndex = index

  const settings = getGameSettings()

  const secretScreen = document.getElementById("secretScreen")
  const secretCard = document.querySelector(".secretCard")
  const secretName = document.getElementById("secretName")
  const secretCategory = document.getElementById("secretCategory")
  const secretWord = document.getElementById("secretWord")

  if (!secretScreen || !secretName || !secretCategory || !secretWord) return

  if (secretCard) {
    secretCard.classList.remove("isImposter")
  }

  secretName.textContent = player.name
  secretWord.classList.remove("imposterWord")

  if (player.isImposter) {
    if (secretCard) {
      secretCard.classList.add("isImposter")
    }

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
  }

  secretScreen.classList.remove("closing")
  secretScreen.classList.add("show")
  secretScreen.style.display = "flex"
}

function hideSecret() {
  const secretScreen = document.getElementById("secretScreen")
  const secretCard = document.querySelector(".secretCard")

  if (currentViewedIndex !== null) {
    players[currentViewedIndex].viewed = true
  }

  currentViewedIndex = null

  if (!secretScreen) {
    renderCards()
    updateCounter()
    updateGameButtons()
    saveGameState()
    return
  }

  secretScreen.classList.remove("show")
  secretScreen.classList.add("closing")

  setTimeout(() => {
  secretScreen.classList.remove("closing")
  secretScreen.style.display = "none"

  if (secretCard) {
    secretCard.classList.remove("isImposter")
  }

  renderCards()
  updateCounter()
  updateGameButtons()
  saveGameState()
}, 120)
}

/* =========================
   القرعة بدون تكرار
========================= */

function shuffleNames(names) {
  const copy = [...names]

  for (let i = copy.length - 1; i > 0; i--) {
    const randomIndex = Math.floor(Math.random() * (i + 1))
    const temp = copy[i]
    copy[i] = copy[randomIndex]
    copy[randomIndex] = temp
  }

  return copy
}

function getNextTurnWinner() {
  const activeNames = players.map((player) => player.name)
  const imposterName = players[imposterIndex]?.name || ""

  const allowedNames = activeNames.filter((name) => {
    return name !== imposterName
  })

  turnDrawPool = turnDrawPool.filter((name) => {
    return allowedNames.includes(name)
  })

  if (turnDrawPool.length === 0) {
    turnDrawPool = shuffleNames(allowedNames)
  }

  const winnerName = turnDrawPool.shift()
  lastTurnWinnerName = winnerName

  return winnerName
}

function startTurnDraw() {
  if (!players.length || turnDrawRunning || turnDrawDone) return

  const screen = document.getElementById("turnDrawScreen")
  const nameBox = document.getElementById("turnDrawName")
  const hint = document.getElementById("turnDrawHint")
  const closeBtn = document.getElementById("closeTurnDrawBtn")

  if (!screen || !nameBox || !hint || !closeBtn) return

  turnDrawRunning = true

  screen.classList.remove("closing")
  screen.classList.add("show")
  screen.style.display = "flex"

  nameBox.classList.remove("winner")
  nameBox.classList.remove("spinning")
  nameBox.textContent = "جاهز؟"

  closeBtn.classList.add("hidden")
  hint.textContent = "جاري اختيار اللاعب..."

  let index = 0
  let loops = 0

  if (turnDrawTimer) {
    clearInterval(turnDrawTimer)
  }

  const winnerName = getNextTurnWinner()

  turnDrawTimer = setInterval(() => {
    nameBox.classList.remove("spinning")
    void nameBox.offsetWidth
    nameBox.classList.add("spinning")

    nameBox.textContent = players[index].name

    index++

    if (index >= players.length) {
      index = 0
      loops++
    }

    if (loops >= 5) {
      clearInterval(turnDrawTimer)
      turnDrawTimer = null

      setTimeout(() => {
        nameBox.classList.remove("spinning")
        nameBox.textContent = winnerName
        nameBox.classList.add("winner")
        hint.textContent = "الدور عليه"
        closeBtn.classList.remove("hidden")
        turnDrawRunning = false
        turnDrawDone = true
        saveGameState()
      }, 260)
    }
  }, 70)
}

function closeTurnDrawScreen() {
  if (turnDrawRunning) return

  const screen = document.getElementById("turnDrawScreen")
  const turnBox = document.getElementById("turnBox")
  const drawTurnBtn = document.getElementById("drawTurnBtn")
  const startVotingBtn = document.getElementById("startVotingBtn")
  const gameStageTitle = document.getElementById("gameStageTitle")

  const turnBoxIcon = document.querySelector(".turnBoxIcon")
  const turnBoxTitle = document.querySelector(".turnBoxText h3")
  const turnBoxText = document.querySelector(".turnBoxText p")

  if (screen) {
    screen.classList.remove("show")
    screen.classList.add("closing")

    setTimeout(() => {
      screen.classList.remove("closing")
      screen.style.display = "none"
    }, 180)
  }

  if (turnDrawDone) {
    if (turnBox) {
      turnBox.classList.remove("hidden", "drawStage")
      turnBox.classList.add("voteReadyStage")
    }

    if (turnBoxIcon) turnBoxIcon.textContent = "🗳️"
    if (turnBoxTitle) turnBoxTitle.textContent = "جاهزين للتصويت؟"
    if (turnBoxText) {
      turnBoxText.textContent = lastTurnWinnerName
        ? `بدأ الدور على ${lastTurnWinnerName}. بعد النقاش، ابدأوا التصويت لاختيار الإمبوستر.`
        : "بعد النقاش، ابدأوا التصويت لاختيار الإمبوستر."
    }

    if (drawTurnBtn) drawTurnBtn.classList.add("hidden")

    if (startVotingBtn) {
      startVotingBtn.classList.remove("hidden")
      startVotingBtn.textContent = "بدء التصويت"
    }

    if (gameStageTitle) gameStageTitle.textContent = "مرحلة ما قبل التصويت"
  }

  saveGameState()
}

function startVotingStage() {
  const settings = getGameSettings()
  const startVotingBtn = document.getElementById("startVotingBtn")
  const turnBox = document.getElementById("turnBox")
  const gameStageTitle = document.getElementById("gameStageTitle")

  votingStarted = true

  if (startVotingBtn) startVotingBtn.classList.add("hidden")

  if (turnBox) {
    turnBox.classList.add("stageLeaving")

    setTimeout(() => {
      turnBox.classList.add("hidden")
      turnBox.classList.remove("voteReadyStage", "drawStage", "stageLeaving")

      if (settings.enableVoting) {
        if (gameStageTitle) gameStageTitle.textContent = "مرحلة التصويت"
        currentVoteIndex = 0
        renderCurrentVote()
      } else {
        if (gameStageTitle) gameStageTitle.textContent = "جاهز لكشف الإمبوستر"
        updateGameButtons()
      }

      saveGameState()
    }, 260)

    return
  }

  if (settings.enableVoting) {
    if (gameStageTitle) gameStageTitle.textContent = "مرحلة التصويت"
    currentVoteIndex = 0
    renderCurrentVote()
  } else {
    if (gameStageTitle) gameStageTitle.textContent = "جاهز لكشف الإمبوستر"
    updateGameButtons()
  }

  saveGameState()
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
  turnDrawDone = false

  if (screen) {
    screen.classList.remove("show")
    screen.classList.remove("closing")
    screen.style.display = "none"
  }

  if (nameBox) {
    nameBox.textContent = "جاهز؟"
    nameBox.classList.remove("winner")
    nameBox.classList.remove("spinning")
  }

  if (hint) hint.textContent = "انتظر حتى تتوقف القرعة"
  if (closeBtn) closeBtn.classList.add("hidden")
}

/* =========================
   أزرار الجولة
========================= */

function updateGameButtons() {
  const revealBtn = document.getElementById("revealBtn")
  const newRoundBtn = document.getElementById("newRoundBtn")
  const resetBtn = document.getElementById("resetBtn")
  const scoreBtn = document.getElementById("scoreBtn")

  if (!revealBtn || !newRoundBtn || !resetBtn || !scoreBtn) return

  const allViewed = areAllCardsViewed()
  const votingComplete = isVotingComplete()

  revealBtn.classList.add("hidden")
  newRoundBtn.classList.add("hidden")
  resetBtn.classList.add("hidden")
  scoreBtn.classList.add("hidden")

  if (!imposterRevealed && allViewed && turnDrawDone && votingStarted && votingComplete) {
    revealBtn.classList.remove("hidden")
  }

  if (imposterRevealed) {
    scoreBtn.classList.remove("hidden")
    newRoundBtn.classList.remove("hidden")
    resetBtn.classList.remove("hidden")
  }
}

function revealImposter() {
  if (!players.length || imposterIndex === null) return

  const imposter = players[imposterIndex]
  const scoreResultText = calculateRoundScores()

  imposterRevealed = true

  const gameStageTitle = document.getElementById("gameStageTitle")
  const votingPanel = document.getElementById("votingPanel")
  const revealScreen = document.getElementById("revealScreen")
  const revealName = document.getElementById("revealName")
  const revealDetails = document.getElementById("revealDetails")

  if (gameStageTitle) gameStageTitle.textContent = "تم كشف الإمبوستر"
  if (votingPanel) votingPanel.classList.add("hidden")

  if (revealName) {
    revealName.textContent = imposter.name
  }

  if (revealDetails) {
    revealDetails.innerHTML = `
      الفئة: ${selectedItem.category}
      <br>
      الكلمة: ${selectedItem.word}
      <br>
      <span style="color:#7F2020">${scoreResultText}</span>
    `
  }

  if (revealScreen) {
    revealScreen.classList.remove("closing")
    revealScreen.classList.add("show")
    revealScreen.style.display = "flex"
  }

  renderScoreBoard()
  updateGameButtons()
  saveGameState()
}

function closeReveal() {
  const revealScreen = document.getElementById("revealScreen")

  if (!revealScreen) return

  revealScreen.classList.remove("show")
  revealScreen.classList.add("closing")

  setTimeout(() => {
    revealScreen.classList.remove("closing")
    revealScreen.style.display = "none"
  }, 180)
}

function newRound() {
  if (players.length === 0) return

  const names = players.map((player) => player.name)

  closeReveal()
  closeScoreBoard()
  startNewRoundWithCurrentPlayers(names)
  saveGameState()
}

function resetGame() {
  players = []
  votes = {}
  scores = {}
  imposterIndex = null
  currentViewedIndex = null
  selectedItem = null
  imposterRevealed = false
  roundScored = false
  currentVoteIndex = 0
  turnDrawDone = false
  votingStarted = false
  turnDrawPool = []
  lastTurnWinnerName = ""

  // يرجع الأسماء الأساسية فقط
  const defaultNames = getDefaultFixedPlayers()
  saveFixedPlayers(defaultNames)

  const secretScreen = document.getElementById("secretScreen")
  const revealScreen = document.getElementById("revealScreen")
  const votingPanel = document.getElementById("votingPanel")
  const cardsGrid = document.getElementById("cardsGrid")
  const turnBox = document.getElementById("turnBox")
  const playersCounter = document.querySelector(".playersCounter")
  const gameStageTitle = document.getElementById("gameStageTitle")

  if (secretScreen) secretScreen.style.display = "none"
  if (revealScreen) revealScreen.style.display = "none"
  if (votingPanel) votingPanel.classList.add("hidden")
  if (cardsGrid) cardsGrid.classList.remove("hidden")
  if (turnBox) turnBox.classList.add("hidden")
  if (playersCounter) playersCounter.classList.remove("hidden")
  if (gameStageTitle) gameStageTitle.textContent = "اختار اسمك وشوف بطاقتك بسرية"

  closeScoreBoard()
  resetTurnBox()
  clearGameState()
  renderFixedPlayers()
  renderScoreBoard()
  updateCounter()
  showScreen("setupScreen")
}

/* =========================
   أدوات عامة
========================= */

function updateCounter() {
  const viewed = players.filter((player) => player.viewed).length
  const total = players.length

  const viewedCounter = document.getElementById("viewedCounter")
  const totalCounter = document.getElementById("totalCounter")

  if (viewedCounter) viewedCounter.textContent = viewed
  if (totalCounter) totalCounter.textContent = total
}

function showScreen(screenId) {
  document.querySelectorAll(".screen").forEach((screen) => {
    screen.classList.remove("active")
  })

  const screen = document.getElementById(screenId)

  if (screen) {
    screen.classList.add("active")
  }
}

function showWarning(text) {
  const warning = document.getElementById("warning")

  if (!warning) return

  warning.textContent = text
  warning.style.display = "block"
}

function hideWarning() {
  const warning = document.getElementById("warning")

  if (!warning) return

  warning.textContent = ""
  warning.style.display = "none"
}

function closeAllOverlaysInstant() {
  const overlays = [
    "secretScreen",
    "revealScreen",
    "turnDrawScreen",
    "scoreScreen"
  ]

  overlays.forEach((id) => {
    const el = document.getElementById(id)

    if (!el) return

    el.classList.remove("show", "closing")
    el.style.display = "none"
  })

  const secretCard = document.querySelector(".secretCard")
  if (secretCard) {
    secretCard.classList.remove("isImposter")
  }
}

function lockAppHeight() {
  document.documentElement.style.setProperty(
    "--app-height",
    `${window.innerHeight}px`
  )
}

/* =========================
   تشغيل أولي
========================= */

window.addEventListener("resize", lockAppHeight)
window.addEventListener("orientationchange", lockAppHeight)

lockAppHeight()
renderFixedPlayers()
applyGameSettingsToUI()
loadExcelFromProject()