(function () {
  "use strict";

  var schools = (window.HZ_SCHOOLS || []).filter(function (s) {
    return s && s.name && s.district && s.tier;
  });

  var districts = [];
  schools.forEach(function (s) {
    if (districts.indexOf(s.district) === -1) districts.push(s.district);
  });
  districts.sort();

  var tiers = [1, 2, 3, 4, 5].filter(function (t) {
    return schools.some(function (s) { return s.tier === t; });
  });

  var REGION_WORDS = ["上城", "拱墅", "西湖", "滨江", "钱塘", "萧山", "余杭", "临平", "富阳", "临安", "桐庐", "淳安", "建德"];
  var GEO_HEADS = ["湾", "河", "湖", "山", "江", "岛", "塘", "街道", "镇", "乡"];
  var SCHOOL_HEADS = ["中学", "学校", "实验", "初级", "第", "一中", "二中", "三中", "四中", "五中", "六中", "十中", "外国语", "教育集团"];
  var GENERIC_NAMES = ["中学", "学校", "实验中学", "实验学校", "实验外国语学校", "外国语学校", "第一中学", "第二中学", "第三中学", "第四中学", "第五中学", "教育集团", "一中实验"];

  function sanitizeName(raw) {
    var s = String(raw || "").trim();
    if (!s) return s;
    if (s.indexOf("杭州市") === 0) {
      var tailCity = s.slice(3);
      if (GENERIC_NAMES.indexOf(tailCity) > -1) {
        return String(raw || "").trim();
      }
      s = tailCity;
    } else if (s.indexOf("杭州") === 0) {
      var tail = s.slice(2);
      if (REGION_WORDS.some(function (r) { return tail.indexOf(r) === 0; })) {
        s = tail;
      }
    }
    var keepOriginal = false;
    REGION_WORDS.forEach(function (r) {
      ["区", "县", "市"].forEach(function (suf) {
        var token = r + suf;
        if (s.indexOf(token) > -1) {
          var candidate = s.split(token).join("");
          if (GENERIC_NAMES.indexOf(candidate) > -1) {
            keepOriginal = true;
          } else {
            s = candidate;
          }
        }
      });
    });
    if (keepOriginal) return String(raw || "").trim();
    for (var i = 0; i < REGION_WORDS.length; i++) {
      var word = REGION_WORDS[i];
      if (s.indexOf(word) !== 0) continue;
      var rest = s.slice(word.length);
      var geoHead = GEO_HEADS.some(function (g) { return rest.indexOf(g) === 0; });
      if (!geoHead && rest && GENERIC_NAMES.indexOf(rest) === -1) {
        s = rest;
        break;
      }
    }
    s = s.replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
    return s || String(raw || "").trim();
  }

  var displayCounts = {};
  schools.forEach(function (s) {
    var d = sanitizeName(s.name);
    displayCounts[d] = (displayCounts[d] || 0) + 1;
  });

  function displayName(school) {
    var d = sanitizeName(school.name);
    return displayCounts[d] > 1 ? (school.name || d) : d;
  }

  var mode = "mix";
  var count = 20;
  var questions = [];
  var current = 0;
  var correctCount = 0;
  var mistakes = [];
  var startedAt = 0;
  var answered = false;

  var $ = function (id) { return document.getElementById(id); };

  var setupView = $("setupView");
  var quizView = $("quizView");
  var resultView = $("resultView");

  function el(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined && text !== null) node.textContent = text;
    return node;
  }

  function shuffle(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = a[i];
      a[i] = a[j];
      a[j] = t;
    }
    return a;
  }

  function pickDistractors(pool, answer, n) {
    var others = pool.filter(function (v) { return v !== answer; });
    return shuffle(others).slice(0, n);
  }

  function buildQuestions() {
    var total = count > 0 ? Math.min(count, schools.length) : schools.length;
    var strongPool = schools.filter(function (s) { return s.tier === 1 || s.tier === 2; });
    var restPool = schools.filter(function (s) {
      return s.tier === null || s.tier === undefined || s.tier >= 3;
    });
    var strongCount = Math.round(total * 0.7);
    var restCount = total - strongCount;

    var pickedStrong = shuffle(strongPool).slice(0, Math.min(strongCount, strongPool.length));
    var pickedRest = shuffle(restPool).slice(0, Math.min(restCount, restPool.length));

    if (pickedStrong.length < strongCount) {
      var need = strongCount - pickedStrong.length;
      var extra = shuffle(restPool.filter(function (s) {
        return pickedRest.indexOf(s) === -1;
      })).slice(0, need);
      pickedStrong = pickedStrong.concat(extra);
    }
    if (pickedRest.length < restCount) {
      var needRest = restCount - pickedRest.length;
      var extraRest = shuffle(strongPool.filter(function (s) {
        return pickedStrong.indexOf(s) === -1;
      })).slice(0, needRest);
      pickedRest = pickedRest.concat(extraRest);
    }

    var picked = shuffle(pickedStrong.concat(pickedRest));
    var qs = picked.map(function (school) {
      var type;
      if (mode === "district") type = "district";
      else if (mode === "tier") type = "tier";
      else type = Math.random() < 0.5 ? "district" : "tier";
      if (type === "tier" && !school.tier) type = "district";

      var options;
      var answer;
      var ask;
      if (type === "district") {
        answer = school.district;
        options = shuffle([answer].concat(pickDistractors(districts, answer, 3)));
        ask = "这所学校位于哪个区？";
      } else {
        answer = "第" + school.tier + "梯队";
        var tierPool = tiers.map(function (t) { return "第" + t + "梯队"; });
        options = shuffle([answer].concat(pickDistractors(tierPool, answer, 3)));
        ask = "这所学校属于第几梯队？";
      }
      return {
        school: school,
        type: type,
        ask: ask,
        options: options,
        answer: answer
      };
    });
    return qs;
  }

  function show(view) {
    [setupView, quizView, resultView].forEach(function (v) {
      v.classList.toggle("hidden", v !== view);
    });
  }

  function metaText() {
    return districts.length + " 区 · " + schools.length + " 所";
  }

  function renderQuestion() {
    var q = questions[current];
    answered = false;

    $("qIndex").textContent = "第 " + (current + 1) + " 题 / 共 " + questions.length + " 题";
    $("qScore").textContent = "已答对 " + correctCount;
    $("progressFill").style.width = (current / questions.length * 100) + "%";
    $("nextBtn").classList.add("hidden");

    var box = $("questionBox");
    box.textContent = "";
    box.appendChild(el("div", "q-type", q.type === "district" ? "地区题" : "梯队题"));
    box.appendChild(el("h2", "q-school", displayName(q.school)));

    var note = [];
    if (q.school.nature) note.push(q.school.nature);
    if (q.school.type) note.push(q.school.type);
    if (note.length > 0) box.appendChild(el("p", "q-note", note.join(" · ")));
    box.appendChild(el("p", "q-ask", q.ask));

    var opts = $("optionsBox");
    opts.textContent = "";
    q.options.forEach(function (opt, i) {
      var btn = el("button", "option");
      btn.type = "button";
      btn.appendChild(el("span", "opt-key", String.fromCharCode(65 + i)));
      btn.appendChild(document.createTextNode(opt));
      btn.addEventListener("click", function () { choose(opt, btn); });
      opts.appendChild(btn);
    });

    $("feedbackBox").classList.add("hidden");
    $("feedbackBox").textContent = "";
  }

  function choose(value, btn) {
    if (answered) return;
    answered = true;

    var q = questions[current];
    var buttons = $("optionsBox").querySelectorAll(".option");
    buttons.forEach(function (b) {
      b.disabled = true;
      if (b.textContent.trim() === q.answer) b.classList.add("correct");
      else if (b !== btn) b.classList.add("dim");
    });

    var tierLabel = q.school.tier ? ("第" + q.school.tier + "梯队") : "暂无梯队数据";
    var fb = $("feedbackBox");
    fb.classList.remove("hidden", "right", "wrong");

    if (value === q.answer) {
      correctCount += 1;
      btn.classList.add("correct");
      fb.classList.add("right");
      fb.textContent = "回答正确 · " + q.school.name + "｜" + q.school.district + " · " + tierLabel;
    } else {
      btn.classList.add("wrong");
      fb.classList.add("wrong");
      fb.textContent = "回答错误 · " + q.school.name + "｜" + q.school.district + " · " + tierLabel;
      mistakes.push({
        school: q.school,
        type: q.type,
        answer: value,
        correct: q.answer
      });
    }

    $("qScore").textContent = "已答对 " + correctCount;
    $("progressFill").style.width = ((current + 1) / questions.length * 100) + "%";

    var next = $("nextBtn");
    next.textContent = current === questions.length - 1 ? "查看成绩" : "下一题";
    next.classList.remove("hidden");
  }

  function startQuiz() {
    questions = buildQuestions();
    current = 0;
    correctCount = 0;
    mistakes = [];
    startedAt = Date.now();
    show(quizView);
    renderQuestion();
  }

  function finish() {
    var total = questions.length;
    var pct = total === 0 ? 0 : Math.round(correctCount / total * 100);
    var seconds = Math.max(1, Math.round((Date.now() - startedAt) / 1000));
    var mm = String(Math.floor(seconds / 60)).padStart(2, "0");
    var ss = String(seconds % 60).padStart(2, "0");
    var elapsed = mm + ":" + ss;

    $("scoreNum").textContent = pct;
    $("scoreRing").style.background = "conic-gradient(var(--accent) " + (pct * 3.6) + "deg, #eceae4 0deg)";

    var stats = $("resultStats");
    stats.textContent = "";
    stats.appendChild(el("div", null, "正确 <strong>" + correctCount + " / " + total + "</strong> 题"));
    stats.appendChild(el("div", null, "用时 <strong>" + elapsed + "</strong>"));
    stats.appendChild(el("div", null, "错题 <strong>" + mistakes.length + "</strong> 道"));

    var list = $("mistakeList");
    list.textContent = "";
    if (mistakes.length === 0) {
      list.appendChild(el("div", "mistakes-empty", "全部答对"));
    } else {
      list.appendChild(el("div", "mistakes-title", "错题回顾 " + mistakes.length));
      mistakes.forEach(function (m) {
        var item = el("div", "mistake-item");
        item.appendChild(el("div", "m-school", m.school.name));
        item.appendChild(el("div", null, "你选：" + m.answer));
        var correctLabel = m.school.tier ? ("第" + m.school.tier + "梯队") : "暂无梯队数据";
        item.appendChild(el("div", "m-correct", "正确：" + m.correct + " · " + m.school.district + " · " + correctLabel));
        list.appendChild(item);
      });
    }

    show(resultView);
  }

  $("bankMeta").textContent = metaText();
  $("bankSummary").textContent = "题库覆盖杭州 " + districts.length + " 个区、" + schools.length + " 所初中，按表格 T1–T5 梯队出题。";

  document.querySelectorAll("#modeSeg .seg-btn").forEach(function (btn) {
    btn.addEventListener("click", function () {
      document.querySelectorAll("#modeSeg .seg-btn").forEach(function (b) { b.classList.remove("active"); });
      btn.classList.add("active");
      mode = btn.dataset.mode;
    });
  });

  document.querySelectorAll("#countSeg .seg-btn").forEach(function (btn) {
    btn.addEventListener("click", function () {
      document.querySelectorAll("#countSeg .seg-btn").forEach(function (b) { b.classList.remove("active"); });
      btn.classList.add("active");
      count = parseInt(btn.dataset.count, 10);
    });
  });

  $("startBtn").addEventListener("click", startQuiz);
  $("quitBtn").addEventListener("click", function () { show(setupView); });
  $("nextBtn").addEventListener("click", function () {
    if (current === questions.length - 1) finish();
    else {
      current += 1;
      renderQuestion();
    }
  });
  $("retryBtn").addEventListener("click", startQuiz);
  $("settingsBtn").addEventListener("click", function () { show(setupView); });

  document.addEventListener("keydown", function (e) {
    if (quizView.classList.contains("hidden")) return;
    if (answered && (e.key === "Enter" || e.key === " ")) {
      e.preventDefault();
      $("nextBtn").click();
      return;
    }
    var idx = ["1", "2", "3", "4"].indexOf(e.key);
    if (!answered && idx > -1) {
      var btns = $("optionsBox").querySelectorAll(".option");
      if (btns[idx]) btns[idx].click();
    }
  });
})();
