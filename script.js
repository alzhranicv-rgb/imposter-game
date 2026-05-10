let wordsList = []
let selectedItem = null

let players = []
let imposterIndex = null
let currentViewedIndex = null

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

    if (wordsList.length === 0) {
      fileInfo.textContent = "لم يتم تحميل كلمات"
      showWarning("ملف Excel لا يحتوي على كلمات واضحة")
      return
    }

    fileInfo.textContent = `تم تحميل ${wordsList.length} كلمة تلقائيًا`
    hideWarning()

    console.log("الكلمات المحملة:", wordsList)

  } catch (error) {
    fileInfo.textContent = "تعذر تحميل ملف الكلمات"
    showWarning("تأكد أن ملف categorized_words.xlsx موجود بجانب index.html وأنك مشغل المشروع عبر Live Server")
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

function removeDuplicateWords(items) {
  const unique = []
  const seen = new Set()

  items.forEach((item) => {
    const key = `${item.category}-${item.word}`

    if (seen.has(key)) return

    seen.add(key)
    unique.push(item)
  })

  return unique
}

function getRandomItem() {
  const randomIndex = Math.floor(Math.random() * wordsList.length)
  return wordsList[randomIndex]
}

function createNameInputs() {
  const countInput = document.getElementById("playerCount")
  const playersInputs = document.getElementById("playersInputs")

  const count = Number(countInput.value)

  hideWarning()
  playersInputs.innerHTML = ""

  if (!count || count < 3) {
    showWarning("لازم يكون عدد اللاعبين 3 أو أكثر")
    return
  }

  if (count > 20) {
    showWarning("الحد الأعلى 20 لاعب")
    countInput.value = 20
    return
  }

  for (let i = 0; i < count; i++) {
    const input = document.createElement("input")

    input.type = "text"
    input.placeholder = `اسم اللاعب ${i + 1}`
    input.id = `playerName${i}`
    input.autocomplete = "off"

    playersInputs.appendChild(input)
  }
}

function startGame() {
  const count = Number(document.getElementById("playerCount").value)

  if (wordsList.length === 0) {
    showWarning("ملف الكلمات لم يتم تحميله بعد")
    return
  }

  selectedItem = getRandomItem()
  players = []

  for (let i = 0; i < count; i++) {
    const input = document.getElementById(`playerName${i}`)
    const name = input ? input.value.trim() : ""

    if (!name) {
      showWarning(`اكتب اسم اللاعب رقم ${i + 1}`)
      return
    }

    players.push({
      name,
      category: selectedItem.category,
      word: selectedItem.word,
      description: selectedItem.description,
      isImposter: false,
      viewed: false
    })
  }

  imposterIndex = Math.floor(Math.random() * players.length)

  players[imposterIndex].isImposter = true
  players[imposterIndex].word = ""

  renderCards()
  updateCounter()

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

  const secretName = document.getElementById("secretName")
  const secretCategory = document.getElementById("secretCategory")
  const secretWord = document.getElementById("secretWord")

  secretName.textContent = player.name
  secretCategory.textContent = `الفئة: ${player.category}`

  if (player.isImposter) {
    secretWord.textContent = "أنت الإمبوستر"
    secretWord.classList.add("imposterWord")
  } else {
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
}

function closeReveal() {
  document.getElementById("revealScreen").style.display = "none"
}

function newRound() {
  if (players.length === 0) return

  selectedItem = getRandomItem()

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
}

function resetGame() {
  players = []
  imposterIndex = null
  currentViewedIndex = null
  selectedItem = null

  document.getElementById("secretScreen").style.display = "none"
  document.getElementById("revealScreen").style.display = "none"

  showScreen("setupScreen")
}

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

createNameInputs()
loadExcelFromProject()