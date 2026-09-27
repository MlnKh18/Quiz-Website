/**
 * ==========================================================================
 * EVALUASI PEMBELAJARAN IPAS KELAS IV SD - JAVASCRIPT ENGINE
 * Bab 7 Topik C: Sikapku Terhadap Keberagaman Budaya
 * Fitur: Form Siswa Lengkap, Sistem Database Lokal & Ekspor Excel (.xlsx)
 * ==========================================================================
 */

// --- 1. DATASET 15 BUTIR SOAL ---
const QUIZ_DATA = [
  {
    id: 1,
    category: "IPAS KELAS IV • BENTUK KEBERAGAMAN",
    type: "multiple_choice",
    question: "Apa yang dimaksud dengan keberagaman budaya di Indonesia?",
    options: [
      "Perbedaan jumlah kekayaan dan uang jajan yang dimiliki setiap keluarga",
      "Macam-macam tradisi, bahasa, pakaian, tarian, dan kesenian yang dimiliki suku bangsa di Indonesia",
      "Perbedaan jenis pekerjaan orang tua siswa di lingkungan sekolah",
      "Perbedaan jenis kendaraan dan gedung bertingkat di perkotaan"
    ],
    answer: "Macam-macam tradisi, bahasa, pakaian, tarian, dan kesenian yang dimiliki suku bangsa di Indonesia",
    explanation: "Keberagaman budaya adalah variasi adat istiadat, bahasa daerah, kesenian, pakaian, dan tradisi yang diwariskan turun-temurun oleh berbagai suku bangsa di Nusantara.",
    score: 10
  },
  {
    id: 2,
    category: "IPAS KELAS IV • BENTUK KEBERAGAMAN",
    type: "multiple_choice",
    question: "Di bawah ini yang merupakan contoh wujud keberagaman budaya asli bangsa Indonesia adalah...",
    options: [
      "Mobil listrik, kereta cepat, dan pesawat terbang",
      "Rumah adat, pakaian daerah, tarian tradisional, dan alat musik daerah",
      "Makanan cepat saji (fast food) dan minuman kaleng impor",
      "Gedung bioskop, jalan tol, dan pusat perbelanjaan"
    ],
    answer: "Rumah adat, pakaian daerah, tarian tradisional, dan alat musik daerah",
    explanation: "Rumah adat, pakaian daerah, tarian, dan alat musik daerah adalah bukti nyata warisan bentuk keragaman budaya asli Nusantara.",
    score: 10
  },
  {
    id: 3,
    category: "IPAS KELAS IV • KEARIFAN LOKAL SEKITAR",
    type: "matching",
    question: "Pasangkan contoh warisan budaya di Kolom Kiri dengan daerah asalnya di Kolom Kanan:",
    pairs: [
      { left: "Kesenian Reog", right: "Ponorogo (Jawa Timur)" },
      { left: "Makanan Tradisional Rendang", right: "Padang (Sumatra Barat)" },
      { left: "Boneka Raksasa Ondel-ondel", right: "Betawi (DKI Jakarta)" }
    ],
    answer: {
      "Kesenian Reog": "Ponorogo (Jawa Timur)",
      "Makanan Tradisional Rendang": "Padang (Sumatra Barat)",
      "Boneka Raksasa Ondel-ondel": "Betawi (DKI Jakarta)"
    },
    explanation: "Reog berasal dari Ponorogo (Jawa Timur), Rendang dari Minangkabau/Padang (Sumatra Barat), dan Ondel-ondel dari Betawi (DKI Jakarta).",
    score: 10
  },
  {
    id: 4,
    category: "IPAS KELAS IV • BENTUK KEBERAGAMAN",
    type: "matching",
    question: "Hubungkan jenis budaya di Kolom Kiri dengan contoh nyatanya di Kolom Kanan:",
    pairs: [
      { left: "Alat Musik Tradisional", right: "Angklung dan Sasando" },
      { left: "Pakaian Adat", right: "Baju Bodo dan Kain Ulos" },
      { left: "Senjata Tradisional", right: "Keris dan Rencong" }
    ],
    answer: {
      "Alat Musik Tradisional": "Angklung dan Sasando",
      "Pakaian Adat": "Baju Bodo dan Kain Ulos",
      "Senjata Tradisional": "Keris dan Rencong"
    },
    explanation: "Angklung dan Sasando adalah alat musik tradisional; Baju Bodo dan Ulos adalah pakaian adat; Keris dan Rencong adalah senjata tradisional.",
    score: 10
  },
  {
    id: 5,
    category: "IPAS KELAS IV • KEKAYAAN BANGSA",
    type: "multiple_choice",
    question: "Mengapa keberagaman budaya yang kita miliki harus dibanggakan dan dijaga bersama?",
    options: [
      "Karena keberagaman budaya membuat Indonesia mudah terpecah belah",
      "Karena keberagaman budaya merupakan warisan luhur yang unik, indah, dan memperkaya identitas bangsa",
      "Karena keberagaman budaya bisa kita gunakan untuk menyombongkan diri kepada bangsa lain",
      "Karena keberagaman budaya membuat masyarakat tidak perlu saling mengenal"
    ],
    answer: "Karena keberagaman budaya merupakan warisan luhur yang unik, indah, dan memperkaya identitas bangsa",
    explanation: "Keberagaman budaya adalah anugerah Tuhan yang membuat bangsa Indonesia memiliki identitas unik, indah, dan saling melengkapi.",
    score: 10
  },
  {
    id: 6,
    category: "IPAS KELAS IV • MANFAAT KEBERAGAMAN",
    type: "multiple_choice",
    question: "Apa manfaat positif dari adanya keberagaman budaya bagi kehidupan masyarakat Indonesia?",
    options: [
      "Menumbuhkan sikap toleransi, mempererat persatuan, dan menarik wisatawan",
      "Menciptakan persaingan yang tidak sehat antarwarga desa",
      "Membuat masyarakat hanya mau bergaul dengan orang yang satu suku",
      "Menyebabkan perbedaan pendapat yang berujung pada pertengkaran"
    ],
    answer: "Menumbuhkan sikap toleransi, mempererat persatuan, dan menarik wisatawan",
    explanation: "Keberagaman budaya melatih masyarakat untuk saling menghormati (toleransi), memperkokoh persaudaraan, serta memperkaya daya tarik wisata.",
    score: 10
  },
  {
    id: 7,
    category: "IPAS KELAS IV • SIKAP BANGGA BUDAYA",
    type: "true_false",
    question: "Mengenakan pakaian batik bermotif daerah pada peringatan hari besar di sekolah dengan senang dan percaya diri merupakan wujud sikap bangga terhadap budaya Indonesia.",
    options: ["Benar", "Salah"],
    answer: "Benar",
    explanation: "Memakai batik dengan senang dan percaya diri membuktikan kecintaan serta kebanggaan kita terhadap warisan seni budaya bangsa sendiri.",
    score: 10
  },
  {
    id: 8,
    category: "IPAS KELAS IV • MENGHARGAI PERBEDAAN",
    type: "multiple_choice",
    question: "Di kelas IV, ada siswa baru bernama Dedi yang berbicara dengan dialek dan logat daerah asalnya yang kental. Sikap yang paling tepat kamu tunjukkan adalah...",
    options: [
      "Menertawakan logat bicaranya bersama teman-teman sekelas",
      "Menghargai cara bicaranya dan tetap mengajaknya mengobrol dengan ramah",
      "Meminta guru agar Dedi tidak diperbolehkan berbicara di kelas",
      "Menjauhi Dedi karena menganggap gaya bicaranya aneh"
    ],
    answer: "Menghargai cara bicaranya dan tetap mengajaknya mengobrol dengan ramah",
    explanation: "Setiap daerah memiliki dialek/logat bahasa yang khas. Kita wajib menghargai perbedaan tersebut dengan tetap bersikap ramah dan bersahabat.",
    score: 10
  },
  {
    id: 9,
    category: "IPAS KELAS IV • IDENTIFIKASI PERILAKU",
    type: "true_false",
    question: "Mengejek gerakan tarian tradisional daerah teman merupakan contoh perilaku yang sesuai dalam menghargai keberagaman budaya.",
    options: ["Benar", "Salah"],
    answer: "Salah",
    explanation: "Mengejek tarian daerah orang lain adalah tindakan tidak terpuji yang dapat menyinggung perasaan teman dan merusak kerukunan.",
    score: 10
  },
  {
    id: 10,
    category: "IPAS KELAS IV • KETERBUKAAN BUDAYA",
    type: "multiple_choice",
    question: "Manakah tindakan di bawah ini yang menunjukkan adanya kemauan dan rasa ingin mengenal budaya dari daerah lain?",
    options: [
      "Membaca buku cerita rakyat Nusantara dan antusias mencicipi kuliner khas daerah lain",
      "Menutup mata dan telinga saat ada pertunjukan kesenian daerah lain",
      "Hanya mau mengonsumsi makanan yang berasal dari daerah sukunya sendiri",
      "Menolak berkunjung ke stan pameran budaya daerah lain saat festival sekolah"
    ],
    answer: "Membaca buku cerita rakyat Nusantara dan antusias mencicipi kuliner khas daerah lain",
    explanation: "Membaca cerita rakyat dan mencoba mencicipi kuliner khas suku lain mencerminkan sikap terbuka untuk mengenal kekayaan Nusantara.",
    score: 10
  },
  {
    id: 11,
    category: "IPAS KELAS IV • KERJA SAMA LINTAS BUDAYA",
    type: "multiple_choice",
    question: "Situasi Belajar: Kelompok belajarmu terdiri dari siswa yang berasal dari suku Jawa, Sunda, Batak, dan Bali. Agar tugas kolase selesai dengan baik, sikap apa yang harus dilakukan?",
    options: [
      "Membagi tugas secara adil serta saling mendengarkan ide kreasi dari setiap anggota",
      "Menyerahkan semua pekerjaan hanya kepada teman yang dianggap paling pintar menggambar",
      "Hanya mau berdiskusi dengan teman yang sukunya sama",
      "Memaksakan rumah adat dari daerah masing-masing tanpa kompromi"
    ],
    answer: "Membagi tugas secara adil serta saling mendengarkan ide kreasi dari setiap anggota",
    explanation: "Kerja sama yang baik tercapai apabila setiap anggota kelompok saling menghargai pendapat, berbagi tugas secara adil, dan gotong royong.",
    score: 10
  },
  {
    id: 12,
    category: "IPAS KELAS IV • TANGGUNG JAWAB SISWA",
    type: "multiple_choice",
    question: "Sebagai seorang murid kelas IV SD, apa contoh tanggung jawabmu dalam menjaga kerukunan dan keberagaman budaya di sekolah?",
    options: [
      "Memilih-milih teman bermain berdasarkan daerah asalnya saja",
      "Bersikap rukun, saling menghormati, dan tidak membeda-bedakan teman saat belajar maupun bermain",
      "Merasa adat tradisi daerah sendiri adalah yang paling unggul di antara yang lain",
      "Melarang teman menampilkan tarian daerah asalnya saat acara kelas"
    ],
    answer: "Bersikap rukun, saling menghormati, dan tidak membeda-bedakan teman saat belajar maupun bermain",
    explanation: "Tanggung jawab utama siswa adalah memelihara kerukunan, belajar bersama secara adil, dan tidak mendiskriminasi latar belakang teman.",
    score: 10
  },
  {
    id: 13,
    category: "IPAS KELAS IV • PELESTARIAN BUDAYA",
    type: "multiple_choice",
    question: "Di bawah ini yang merupakan tindakan sederhana seorang siswa kelas IV untuk ikut melestarikan budaya daerah di sekolah adalah...",
    options: [
      "Mengikuti latihan ekstrakurikuler seni tari tradisional atau musik daerah dengan tekun",
      "Malu saat diminta menyanyikan lagu daerah di depan kelas",
      "Menolak belajar muatan lokal bahasa daerah yang diajarkan oleh guru",
      "Hanya bermain game daring dan melupakan seluruh permainan tradisional"
    ],
    answer: "Mengikuti latihan ekstrakurikuler seni tari tradisional atau musik daerah dengan tekun",
    explanation: "Aktif mempelajari kesenian tradisional di sekolah merupakan wujud nyata kontribusi siswa dalam melestarikan budaya bangsa agar tidak punah.",
    score: 10
  },
  {
    id: 14,
    category: "IPAS KELAS IV • AKSI & TUJUAN PELESTARIAN",
    type: "matching",
    question: "Pasangkan tindakan pelestarian budaya di Kolom Kiri dengan tujuan pelestariannya di Kolom Kanan:",
    pairs: [
      { left: "Mempelajari tari daerah di sanggar", right: "Menjaga seni pertunjukan tradisional tetap lestari" },
      { left: "Mengadakan festival jajanan tradisional", right: "Mengenalkan kelezatan kuliner khas daerah" },
      { left: "Berkunjung ke rumah adat di museum", right: "Mengenal sejarah dan arsitektur rumah adat" }
    ],
    answer: {
      "Mempelajari tari daerah di sanggar": "Menjaga seni pertunjukan tradisional tetap lestari",
      "Mengadakan festival jajanan tradisional": "Mengenalkan kelezatan kuliner khas daerah",
      "Berkunjung ke rumah adat di museum": "Mengenal sejarah dan arsitektur rumah adat"
    },
    explanation: "Belajar tari menjaga seni pertunjukan tetap hidup; festival makanan mengenalkan kuliner daerah; kunjungan museum memperluas wawasan arsitektur.",
    score: 10
  },
  {
    id: 15,
    category: "IPAS KELAS IV • ANALISIS KASUS SEKOLAH",
    type: "multiple_select",
    question: "Kasus: Pada acara Pentas Seni Sekolah, Riko sempat berbisik mengejek hiasan kepala pakaian adat Wayan. Guru mengingatkan semboyan Bhinneka Tunggal Ika. Manakah dua (2) tindakan yang tepat dilakukan?",
    options: [
      "Riko menyadari kesalahan, meminta maaf, dan mengapresiasi pakaian adat Wayan",
      "Wayan membalas dengan merusak hiasan pakaian adat milik Riko",
      "Seluruh siswa saling memberikan tepuk tangan apresiasi kepada seluruh penampil",
      "Meminta pihak sekolah untuk membatalkan pementasan seni budaya"
    ],
    answer: [
      "Riko menyadari kesalahan, meminta maaf, dan mengapresiasi pakaian adat Wayan",
      "Seluruh siswa saling memberikan tepuk tangan apresiasi kepada seluruh penampil"
    ],
    explanation: "Penyelesaian terbaik dalam keberagaman adalah berani meminta maaf, menghormati pakaian adat, dan saling memberikan apresiasi positif.",
    score: 10
  }
];

// --- 2. LOCAL DATABASE & STATE MANAGER ---
const DB_STORAGE_KEY = "IPAS_QUIZ_DATABASE_RECORDS_V1";
const SETTINGS_KEY    = "IPAS_QUIZ_ADMIN_SETTINGS_V1";

function getSettings() {
  try { return JSON.parse(localStorage.getItem(SETTINGS_KEY)) || {}; }
  catch(e) { return {}; }
}

const state = {
  questions: QUIZ_DATA,
  studentName: "",
  studentClass: "Kelas IV-A",
  studentNumber: "1",
  currentIndex: 0,
  score: 0,
  correctCount: 0,
  wrongCount: 0,
  isAnswered: false,
  userResponses: [],
  lastSavedRecord: null
};

// --- 3. DOM ELEMENTS CACHE ---
const dom = {
  // Screens
  startScreen: document.getElementById("start-screen"),
  quizScreen: document.getElementById("quiz-screen"),
  resultScreen: document.getElementById("result-screen"),
  reviewModal: document.getElementById("review-modal"),

  // Start Form
  studentNameInput: document.getElementById("student-name-input"),
  studentClassInput: document.getElementById("student-class-input"),
  studentNumberInput: document.getElementById("student-number-input"),
  btnStartQuiz: document.getElementById("btn-start-quiz"),

  // Quiz Top Nav
  btnQuizBack: document.getElementById("btn-quiz-back"),
  questionCounterLabel: document.getElementById("question-counter-label"),
  quizScorePill: document.getElementById("quiz-score-pill"),
  quizProgressBar: document.getElementById("quiz-progress-bar"),

  // Card Content
  cardCategoryLabel: document.getElementById("card-category-label"),
  questionText: document.getElementById("question-text"),
  answerContainer: document.getElementById("answer-container"),

  // Inline Feedback
  inlineFeedbackBox: document.getElementById("inline-feedback-box"),
  feedbackStatusDot: document.getElementById("feedback-status-dot"),
  feedbackStatusText: document.getElementById("feedback-status-text"),
  feedbackExplanationText: document.getElementById("feedback-explanation-text"),

  // Next Action Button
  btnNextAction: document.getElementById("btn-next-action"),
  btnNextLabel: document.getElementById("btn-next-label"),

  // Result Screen
  resName: document.getElementById("res-name"),
  resClass: document.getElementById("res-class"),
  resultFinalScore: document.getElementById("result-final-score"),
  resultCorrectCount: document.getElementById("result-correct-count"),
  resultWrongCount: document.getElementById("result-wrong-count"),
  resultGrade100: document.getElementById("result-grade-100"),
  resultEvaluationText: document.getElementById("result-evaluation-text"),
  syncStatusBadge: document.getElementById("sync-status-badge"),
  syncDot: document.getElementById("sync-dot"),
  syncStatusText: document.getElementById("sync-status-text"),
  btnOpenReview: document.getElementById("btn-open-review"),
  btnRestartQuiz: document.getElementById("btn-restart-quiz"),

  // Review Modal
  reviewItemsContainer: document.getElementById("review-items-container"),
  btnCloseReview: document.getElementById("btn-close-review"),
  btnDoneReview: document.getElementById("btn-done-review")
};

// --- 4. SVG ICON GENERATORS ---
function getCheckSvg() {
  return `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="20 6 9 17 4 12"></polyline>
    </svg>
  `;
}

function getCrossSvg() {
  return `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round">
      <line x1="18" y1="6" x2="6" y2="18"></line>
      <line x1="6" y1="6" x2="18" y2="6"></line>
    </svg>
  `;
}

// --- 5. INITIALIZATION ---
document.addEventListener("DOMContentLoaded", () => {
  setupEventListeners();
});

function setupEventListeners() {
  // Start Screen Events
  dom.btnStartQuiz.addEventListener("click", handleStartQuiz);
  dom.studentNameInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); handleStartQuiz(); }
  });

  // Quiz Navigation
  dom.btnQuizBack.addEventListener("click", () => {
    if (confirm("Apakah Anda yakin ingin keluar dan membatalkan pengerjaan saat ini?")) {
      switchView("start");
    }
  });

  dom.btnNextAction.addEventListener("click", handleNextAction);

  // Result Actions
  dom.btnRestartQuiz.addEventListener("click", () => {
    dom.studentNameInput.value = "";
    dom.studentNumberInput.value = "";
    switchView("start");
  });

  // Review Modal
  dom.btnOpenReview.addEventListener("click", openReviewModal);
  dom.btnCloseReview.addEventListener("click", closeReviewModal);
  dom.btnDoneReview.addEventListener("click", closeReviewModal);
}

// --- 6. SCREEN ROUTER ---
function switchView(viewName) {
  dom.startScreen.classList.remove("active");
  dom.quizScreen.classList.remove("active");
  dom.resultScreen.classList.remove("active");

  dom.startScreen.classList.add("hidden");
  dom.quizScreen.classList.add("hidden");
  dom.resultScreen.classList.add("hidden");

  if (viewName === "start") {
    dom.startScreen.classList.remove("hidden");
    dom.startScreen.classList.add("active");
  } else if (viewName === "quiz") {
    dom.quizScreen.classList.remove("hidden");
    dom.quizScreen.classList.add("active");
  } else if (viewName === "result") {
    dom.resultScreen.classList.remove("hidden");
    dom.resultScreen.classList.add("active");
  }
}

// --- 7. START QUIZ ---
function handleStartQuiz() {
  const name = dom.studentNameInput.value.trim();
  const className = dom.studentClassInput.value.trim();
  const number = dom.studentNumberInput.value.trim();

  if (!name) {
    dom.studentNameInput.focus();
    dom.studentNameInput.style.borderColor = "#EF4444";
    setTimeout(() => { dom.studentNameInput.style.borderColor = ""; }, 1200);
    return;
  }

  if (!number) {
    dom.studentNumberInput.focus();
    dom.studentNumberInput.style.borderColor = "#EF4444";
    setTimeout(() => { dom.studentNumberInput.style.borderColor = ""; }, 1200);
    return;
  }

  state.studentName = name;
  state.studentClass = className || "Kelas IV-A";
  state.studentNumber = number;
  state.currentIndex = 0;
  state.score = 0;
  state.correctCount = 0;
  state.wrongCount = 0;
  state.userResponses = [];
  state.lastSavedRecord = null;

  updateScorePill();
  switchView("quiz");
  renderQuestion();
}

function updateScorePill() {
  dom.quizScorePill.textContent = `Skor: ${state.score}`;
}

// --- 8. QUESTION RENDERER ---
function renderQuestion() {
  state.isAnswered = false;
  const q = state.questions[state.currentIndex];
  const total = state.questions.length;
  const currentNum = state.currentIndex + 1;

  // Format Counter: "Soal 1 dari 15"
  dom.questionCounterLabel.textContent = `Soal ${currentNum} dari ${total}`;

  // Progress Bar Fill
  const progressPercent = (currentNum / total) * 100;
  dom.quizProgressBar.style.width = `${progressPercent}%`;

  // Set Category & Title
  dom.cardCategoryLabel.textContent = q.category || "IPAS KELAS IV • KEBERAGAMAN BUDAYA";
  dom.questionText.textContent = q.question;

  // Reset Feedback & Next Button State
  dom.inlineFeedbackBox.classList.add("hidden");
  dom.btnNextAction.disabled = true;
  dom.btnNextLabel.textContent = (currentNum === total) ? "Lihat Hasil Nilai" : "Lanjut";

  dom.answerContainer.innerHTML = "";

  if (q.type === "multiple_choice") {
    renderMultipleChoice(q);
  } else if (q.type === "true_false") {
    renderTrueFalse(q);
  } else if (q.type === "matching") {
    renderMatching(q);
  } else if (q.type === "multiple_select") {
    renderMultipleSelect(q);
  }
}

// A. Multiple Choice
function renderMultipleChoice(q) {
  q.options.forEach((optText) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "option-item";

    const textSpan = document.createElement("span");
    textSpan.className = "option-text";
    textSpan.textContent = optText;

    const badgeSpan = document.createElement("span");
    badgeSpan.className = "option-status-badge hidden";

    btn.appendChild(textSpan);
    btn.appendChild(badgeSpan);

    btn.addEventListener("click", () => {
      if (state.isAnswered) return;
      state.isAnswered = true;
      handleSingleAnswerSelection(q, optText, btn);
    });

    dom.answerContainer.appendChild(btn);
  });
}

// B. True / False
function renderTrueFalse(q) {
  q.options.forEach((optText) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "option-item";

    const textSpan = document.createElement("span");
    textSpan.className = "option-text";
    textSpan.textContent = optText;

    const badgeSpan = document.createElement("span");
    badgeSpan.className = "option-status-badge hidden";

    btn.appendChild(textSpan);
    btn.appendChild(badgeSpan);

    btn.addEventListener("click", () => {
      if (state.isAnswered) return;
      state.isAnswered = true;
      handleSingleAnswerSelection(q, optText, btn);
    });

    dom.answerContainer.appendChild(btn);
  });
}

// C. Matching
function renderMatching(q) {
  const wrapper = document.createElement("div");
  wrapper.style.display = "flex";
  wrapper.style.flexDirection = "column";
  wrapper.style.gap = "8px";

  const rightOptions = q.pairs.map(p => p.right).sort(() => Math.random() - 0.5);
  const selects = [];

  q.pairs.forEach((pair, idx) => {
    const row = document.createElement("div");
    row.className = "matching-pair-row";

    const leftCol = document.createElement("div");
    leftCol.className = "matching-pair-label";
    leftCol.textContent = `${idx + 1}. ${pair.left}`;

    const select = document.createElement("select");
    select.className = "matching-select";
    select.dataset.leftKey = pair.left;

    const defaultOpt = document.createElement("option");
    defaultOpt.value = "";
    defaultOpt.textContent = "Pilih Pasangan...";
    select.appendChild(defaultOpt);

    rightOptions.forEach(rText => {
      const opt = document.createElement("option");
      opt.value = rText;
      opt.textContent = rText;
      select.appendChild(opt);
    });

    selects.push(select);
    row.appendChild(leftCol);
    row.appendChild(select);
    wrapper.appendChild(row);
  });

  const confirmBtn = document.createElement("button");
  confirmBtn.type = "button";
  confirmBtn.className = "btn-action-outline";
  confirmBtn.style.marginTop = "8px";
  confirmBtn.textContent = "Kirim Pasangan";

  confirmBtn.addEventListener("click", () => {
    if (state.isAnswered) return;

    const userPairs = {};
    let allSelected = true;

    selects.forEach(sel => {
      if (!sel.value) allSelected = false;
      userPairs[sel.dataset.leftKey] = sel.value;
    });

    if (!allSelected) {
      alert("Silakan pasangkan seluruh pilihan terlebih dahulu.");
      return;
    }

    state.isAnswered = true;
    selects.forEach(sel => sel.disabled = true);
    confirmBtn.disabled = true;

    let isAllCorrect = true;
    for (const key in q.answer) {
      if (userPairs[key] !== q.answer[key]) {
        isAllCorrect = false;
        break;
      }
    }

    evaluateAndShowFeedback(q, isAllCorrect, userPairs);
  });

  wrapper.appendChild(confirmBtn);
  dom.answerContainer.appendChild(wrapper);
}

// D. Multiple Select
function renderMultipleSelect(q) {
  const wrapper = document.createElement("div");
  wrapper.style.display = "flex";
  wrapper.style.flexDirection = "column";
  wrapper.style.gap = "10px";

  const checkboxes = [];

  q.options.forEach((optText) => {
    const item = document.createElement("label");
    item.className = "option-item";
    item.style.cursor = "pointer";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "checkbox-custom";
    cb.value = optText;

    const textSpan = document.createElement("span");
    textSpan.className = "option-text";
    textSpan.textContent = optText;

    item.appendChild(cb);
    item.appendChild(textSpan);

    checkboxes.push(cb);
    wrapper.appendChild(item);
  });

  const confirmBtn = document.createElement("button");
  confirmBtn.type = "button";
  confirmBtn.className = "btn-action-outline";
  confirmBtn.style.marginTop = "8px";
  confirmBtn.textContent = "Kirim Jawaban";

  confirmBtn.addEventListener("click", () => {
    if (state.isAnswered) return;

    const selected = checkboxes.filter(cb => cb.checked).map(cb => cb.value);
    if (selected.length === 0) {
      alert("Pilih minimal satu jawaban.");
      return;
    }

    state.isAnswered = true;
    checkboxes.forEach(cb => cb.disabled = true);
    confirmBtn.disabled = true;

    const isCorrect = (
      selected.length === q.answer.length &&
      selected.every(val => q.answer.includes(val))
    );

    evaluateAndShowFeedback(q, isCorrect, selected);
  });

  wrapper.appendChild(confirmBtn);
  dom.answerContainer.appendChild(wrapper);
}

// --- 9. ANSWER EVALUATION ---
function handleSingleAnswerSelection(question, selectedText, clickedButton) {
  const isCorrect = (selectedText === question.answer);
  const allButtons = dom.answerContainer.querySelectorAll(".option-item");

  allButtons.forEach((btn) => {
    btn.disabled = true;
    const btnText = btn.querySelector(".option-text").textContent;
    const badge = btn.querySelector(".option-status-badge");

    if (btnText === question.answer) {
      btn.classList.add("state-correct");
      badge.innerHTML = getCheckSvg();
      badge.classList.remove("hidden");
    } else if (btn === clickedButton && !isCorrect) {
      btn.classList.add("state-wrong");
      badge.innerHTML = getCrossSvg();
      badge.classList.remove("hidden");
    }
  });

  evaluateAndShowFeedback(question, isCorrect, selectedText);
}

function evaluateAndShowFeedback(question, isCorrect, userAnswer) {
  if (isCorrect) {
    state.score += (question.score || 10);
    state.correctCount++;
  } else {
    state.wrongCount++;
  }

  updateScorePill();

  state.userResponses.push({
    question: question,
    userAnswer: userAnswer,
    isCorrect: isCorrect
  });

  dom.feedbackStatusDot.className = `status-indicator-dot ${isCorrect ? "correct" : "wrong"}`;
  dom.feedbackStatusText.className = `status-indicator-text ${isCorrect ? "correct" : "wrong"}`;
  dom.feedbackStatusText.textContent = isCorrect ? "Jawabanmu Benar!" : "Jawaban Kurang Tepat";
  dom.feedbackExplanationText.textContent = question.explanation;
  dom.inlineFeedbackBox.classList.remove("hidden");

  dom.btnNextAction.disabled = false;
}

// --- 10. NEXT ACTION HANDLER ---
function handleNextAction() {
  if (state.currentIndex < state.questions.length - 1) {
    state.currentIndex++;
    renderQuestion();
  } else {
    showResultScreen();
  }
}

// --- 11. RESULT SCREEN, AUTO DB SAVE & GOOGLE SHEETS SYNC ---
function showResultScreen() {
  switchView("result");

  const grade100 = Math.round((state.score / 150) * 100);

  dom.resName.textContent = state.studentName;
  dom.resClass.textContent = state.studentClass;

  dom.resultFinalScore.textContent = state.score;
  dom.resultCorrectCount.textContent = state.correctCount;
  dom.resultWrongCount.textContent = state.wrongCount;
  dom.resultGrade100.textContent = `${grade100}`;

  let predicate = "Cukup";
  if (state.score >= 140) {
    predicate = "Sangat Baik";
    dom.resultEvaluationText.textContent = `Hebat sekali! Pemahamanmu tentang materi keberagaman budaya sudah sangat baik. Pertahankan terus ya!`;
  } else if (state.score >= 100) {
    predicate = "Baik";
    dom.resultEvaluationText.textContent = `Bagus! Kamu sudah memahami materi keberagaman budaya dengan baik. Tetap rajin belajar ya!`;
  } else {
    predicate = "Cukup";
    dom.resultEvaluationText.textContent = `Tetap semangat! Coba baca dan pelajari lagi materi Bab 7 agar pemahamanmu semakin mantap.`;
  }

  // Auto Save to localStorage
  const record = {
    id: "REC_" + Date.now(),
    timestamp: getFormattedDateTime(),
    name: state.studentName,
    className: state.studentClass,
    number: state.studentNumber,
    correct: state.correctCount,
    wrong: state.wrongCount,
    score: state.score,
    grade100: grade100,
    predicate: predicate
  };

  saveRecordToDatabase(record);
  state.lastSavedRecord = record;

  // Attempt Google Sheets sync
  syncToGoogleSheets(record);
}

// Google Sheets Sync via Apps Script Web App
function syncToGoogleSheets(record) {
  const settings = getSettings();
  const webAppUrl = settings.webAppUrl || "";

  if (!webAppUrl) {
    // No URL configured — show stored-locally state
    setSyncStatus("local", "Tersimpan di perangkat ini");
    return;
  }

  setSyncStatus("syncing", "Mengirim data ke Google Sheets...");

  const payload = {
    nama: record.name,
    kelas: `${record.className} (Absen ${record.number})`,
    absen: record.number,
    skor: record.correct,       // jumlah soal benar
    total: 15,                   // total soal
    nilai: record.grade100,     // nilai skala 100
    skor_mentah: record.score,  // skor mentah /150
    predicate: record.predicate,
    timestamp: record.timestamp,
    jawaban: state.userResponses.reduce((acc, r, i) => {
      acc[`soal_${i + 1}`] = r.isCorrect ? "benar" : "salah";
      return acc;
    }, {})
  };

  fetch(webAppUrl, {
    method: "POST",
    mode: "no-cors",            // Apps Script requires no-cors
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  })
  .then(() => {
    // no-cors gives opaque response; treat as success if no error
    setSyncStatus("success", "Berhasil tersimpan ke Google Sheets");
  })
  .catch((err) => {
    console.warn("Sheets sync failed:", err);
    setSyncStatus("error", "Gagal terhubung ke Google Sheets. Data tersimpan di perangkat.");
  });
}

function setSyncStatus(status, message) {
  const dot = dom.syncDot;
  const text = dom.syncStatusText;
  const badge = dom.syncStatusBadge;

  dot.className = "sync-dot-pulse";
  badge.className = "sync-status-badge";

  if (status === "syncing") {
    dot.classList.add("syncing");
    badge.classList.add("syncing");
  } else if (status === "success") {
    dot.classList.add("success");
    badge.classList.add("success");
  } else if (status === "error") {
    dot.classList.add("error");
    badge.classList.add("error");
  } else {
    dot.classList.add("local");
    badge.classList.add("local");
  }

  text.textContent = message;
}

function getFormattedDateTime() {
  const d = new Date();
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const mins = String(d.getMinutes()).padStart(2, "0");
  return `${day}/${month}/${year} ${hours}:${mins}`;
}

// --- 12. DATABASE CRUD CONTROLLER (LOCAL STORAGE) ---
function getDatabaseRecords() {
  try {
    const raw = localStorage.getItem(DB_STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveRecordToDatabase(record) {
  const records = getDatabaseRecords();
  records.unshift(record);
  localStorage.setItem(DB_STORAGE_KEY, JSON.stringify(records));
}

function getFormattedDateForFilename() {
  const d = new Date();
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  return `${year}${month}${day}`;
}

// --- 15. REVIEW MODAL ---
function openReviewModal() {
  dom.reviewItemsContainer.innerHTML = "";

  state.userResponses.forEach((item, idx) => {
    const q = item.question;
    const isCorr = item.isCorrect;

    const row = document.createElement("div");
    row.className = "review-card-item";

    let userAnsDisplay = "";
    let correctAnsDisplay = "";

    if (q.type === "matching") {
      userAnsDisplay = Object.entries(item.userAnswer).map(([k, v]) => `${k} → ${v}`).join("; ");
      correctAnsDisplay = Object.entries(q.answer).map(([k, v]) => `${k} → ${v}`).join("; ");
    } else if (Array.isArray(item.userAnswer)) {
      userAnsDisplay = item.userAnswer.join(", ");
      correctAnsDisplay = q.answer.join(", ");
    } else {
      userAnsDisplay = item.userAnswer;
      correctAnsDisplay = q.answer;
    }

    row.innerHTML = `
      <div class="review-item-top">
        <span class="review-item-number">Soal ${idx + 1}</span>
        <span class="review-item-badge ${isCorr ? "correct" : "wrong"}">
          ${isCorr ? "Benar (+10)" : "Salah (0)"}
        </span>
      </div>
      <div class="review-question-title">${q.question}</div>
      <div class="review-details">
        <div><strong>Jawaban Kamu:</strong> ${userAnsDisplay}</div>
        <div class="review-ans-correct"><strong>Kunci Jawaban:</strong> ${correctAnsDisplay}</div>
        <div style="margin-top: 6px; color: var(--text-body); background: #F1F5F9; padding: 8px 10px; border-radius: 8px;"><strong>Pembahasan:</strong> ${q.explanation}</div>
      </div>
    `;

    dom.reviewItemsContainer.appendChild(row);
  });

  dom.reviewModal.classList.remove("hidden");
}

function closeReviewModal() {
  dom.reviewModal.classList.add("hidden");
}
