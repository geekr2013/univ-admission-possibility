const YEAR_LABEL = (year) => `${String(year).match(/20\d{2}/)?.[0] || year}학년도 정시`;

const n = (value, fallback = 0) => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};

const positive = (value) => Number.isFinite(Number(value)) && Number(value) > 0;
const avg = (...values) => values.reduce((sum, value) => sum + n(value), 0) / values.length;
const excelRound = (value, digits = 0) => {
  const factor = 10 ** digits;
  return Math.round((n(value) + Number.EPSILON) * factor) / factor;
};
const largest = (values, rank) => [...values].sort((a, b) => b - a)[rank - 1] || 0;
const isScience = (subject) => /[ⅠⅡ]$/.test(String(subject || ""));

function splitRatio(value) {
  const match = String(value ?? "").trim().match(/^(-?[\d.]+)(?:\((-?[\d.]+)\))?$/);
  if (!match) return [0, 0];
  const primary = n(match[1]);
  return [primary, match[2] == null ? primary : n(match[2])];
}

function scoreKey(year, grade, month, subject, scoreGrade) {
  return `${year}|${grade}|${month}|${subject}|${scoreGrade}`;
}

function csatKey(examYear, subject, grade) {
  return `${examYear}|${subject}|${grade}`;
}

export function createExcelEngine(data) {
  const mockIndex = new Map();
  for (const row of data.gradeRanges || []) {
    if (typeof row.year !== "number") continue;
    mockIndex.set(scoreKey(row.year, row.gradeLevel, row.month, row.subject, row.grade), row);
  }

  const csatIndex = new Map();
  for (const row of data.csatGradeRanges || []) {
    csatIndex.set(csatKey(row.examYear, row.subject, row.grade), row);
  }

  function convertedStandard(input, subject, exam, targetExamYear) {
    const grade = Math.max(1, Math.min(9, Math.trunc(n(input.grade, 9))));
    const mock = mockIndex.get(scoreKey(n(exam.year), exam.grade, exam.month, subject, grade));
    const csat = csatIndex.get(csatKey(targetExamYear, subject, grade));
    if (!mock || !csat) return 0;
    const denominator = n(mock.standardMax) - n(mock.standardMin);
    const position = denominator === 0 ? 0 : Math.max(0, Math.min(1, (n(input.standard) - n(mock.standardMin)) / denominator));
    return excelRound(n(csat.standardMin) + position * (n(csat.standardMax) - n(csat.standardMin)), 0);
  }

  function csatMaximum(targetExamYear, subject) {
    return n(csatIndex.get(csatKey(targetExamYear, subject, 1))?.standardMax);
  }

  function conversionColumn(item, subject, targetExamYear) {
    const university = item.university;
    const type = String(item.conversionType || "");
    const science = isScience(subject);
    const suffix = science ? "과탐" : "사탐";
    const year = Number(String(targetExamYear).match(/20\d{2}/)?.[0]);

    if (year === 2024) {
      if (type) {
        if (["광운대학교", "한양대학교"].includes(university)) {
          if (university === "광운대학교" && type === "B" && !science) return `${university}-A`;
          return `${university}-${type}`;
        }
        if (university === "동국대학교") return type === "C" ? `${university}-${type}-${suffix}` : `${university}-${type}`;
        if (university === "성균관대학교") return `${university}-${type}-${suffix}`;
        return `${university}-${type}`;
      }
      if (["경북대학교", "고려대학교", "서울시립대학교", "세종대학교", "아주대학교", "연세대학교_미래", "전북대학교", "고려대학교_세종", "이화여자대학교"].includes(university)) {
        return `${university}-${suffix}`;
      }
    }

    if (year === 2025) {
      if (["고려대학교", "고려대학교_세종", "서울시립대학교", "성균관대학교", "아주대학교", "전북대학교", "경북대학교"].includes(university)) {
        return `${university}-${suffix}`;
      }
      if (university === "가톨릭대학교") {
        if (["의예과", "약학과"].includes(item.major)) return `${university}-약의`;
        if (item.major === "간호학과") return `${university}-간호`;
      }
    }

    if (year === 2026) {
      if (["서강대학교", "성균관대학교", "중앙대학교"].includes(university)) return `${university}-${type}`;
      if (["서울시립대학교", "이화여자대학교", "전북대학교"].includes(university)) return `${university}-${suffix}`;
    }
    return university;
  }

  function inquiryScore(item, input, exam, targetExamYear) {
    let subject = String(input.subject || "").trim();
    if (subject === "사회탐구") subject = "사회·문화";
    if (subject === "과학탐구") subject = "지구과학Ⅰ";
    const metric = String(item.metric || "");
    if (metric === "백분위") return { subject, value: n(input.percentile), maximum: 100 };
    if (metric === "등급") return { subject, value: n(input.grade), maximum: 1 };
    if (metric === "표준점수") {
      const value = item.university === "단국대학교_천안" && /202[45]/.test(targetExamYear)
        ? n(input.percentile)
        : convertedStandard(input, subject, exam, targetExamYear);
      return { subject, value, maximum: csatMaximum(targetExamYear, subject) };
    }
    if (metric === "변환표준점수") {
      const column = conversionColumn(item, subject, targetExamYear);
      const percentile = String(excelRound(n(input.percentile), 0));
      const table = data.conversionTables?.[String(targetExamYear).slice(0, 4)]?.[column] || {};
      return { subject, value: n(table[percentile], n(input.percentile)), maximum: n(table["100"], csatMaximum(targetExamYear, subject)) };
    }
    return { subject, value: n(input.percentile), maximum: 100 };
  }

  function processScores(item, inputs) {
    const targetExamYear = YEAR_LABEL(item.year);
    const exam = inputs.exam;
    const metric = String(item.metric || "");
    const koreanValue = metric === "백분위" ? n(inputs.korean.percentile)
      : metric === "등급" ? n(inputs.korean.grade)
      : convertedStandard(inputs.korean, "국어", exam, targetExamYear);
    const mathValue = metric === "백분위" ? n(inputs.math.percentile)
      : metric === "등급" ? n(inputs.math.grade)
      : convertedStandard(inputs.math, "수학", exam, targetExamYear);
    const inquiry1 = inquiryScore(item, inputs.explore1, exam, targetExamYear);
    const inquiry2 = inquiryScore(item, inputs.explore2, exam, targetExamYear);
    const englishGrade = Math.max(1, Math.min(9, Math.trunc(n(inputs.english.grade, 9))));
    const historyGrade = Math.max(1, Math.min(9, Math.trunc(n(inputs.history.grade, 9))));
    const secondGrade = Math.max(1, Math.min(9, Math.trunc(n(inputs.language2?.grade, 9))));
    const historyPair = splitRatio(item.historyPoints?.[historyGrade - 1]);
    return {
      targetExamYear,
      korean: koreanValue,
      math: mathValue,
      english: n(item.englishPoints?.[englishGrade - 1]),
      history: historyPair[0],
      historyOther: historyPair[1],
      inquiry1: inquiry1.value,
      inquiry2: inquiry2.value,
      inquiry1Subject: inquiry1.subject,
      inquiry2Subject: inquiry2.subject,
      secondLanguage: item.university === "서울대학교" ? n(item.secondLanguagePoints?.[secondGrade - 1]) : 0,
      maxima: [
        csatMaximum(targetExamYear, "국어"),
        csatMaximum(targetExamYear, "수학"),
        n(item.englishPoints?.[0]),
        inquiry1.maximum,
        inquiry2.maximum,
      ],
    };
  }

  function applyBonuses(item, scores, inputs) {
    const bonus = (item.bonusColumns || []).map((value) => n(value));
    const cases = { 2: 0, 3: 0, 4: 0, 7: 0, 8: 0, 9: 0, 10: 0 };
    if (String(item.bonusColumns?.[0] || "") !== "Y") return cases;
    let mathSubject = inputs.math.subject || "";
    if (mathSubject === "수학") mathSubject = "미적분";
    const sci1 = isScience(scores.inquiry1Subject);
    const sci2 = isScience(scores.inquiry2Subject);

    if (mathSubject === "확률과 통계" && bonus[1] > 0) scores.math *= 1 + bonus[1];
    if (mathSubject === "미적분" && bonus[2] > 0) {
      if (item.university === "경상국립대학교") cases[2] = bonus[2];
      else if (item.university === "국립부경대학교") cases[2] = scores.math * (1 + bonus[2]);
      else if (item.university !== "동덕여자대학교" || !(/2025/.test(scores.targetExamYear) && item.major === "자연정보융합학부") || scores.math > scores.korean) scores.math *= 1 + bonus[2];
    }
    if (mathSubject === "기하" && bonus[3] > 0) {
      if (item.university === "경상국립대학교") cases[3] = bonus[3];
      else if (item.university === "국립부경대학교") cases[3] = scores.math * (1 + bonus[3]);
      else if (item.university !== "동덕여자대학교" || !(/2025/.test(scores.targetExamYear) && item.major === "자연정보융합학부") || scores.math > scores.korean) scores.math *= 1 + bonus[3];
    }

    if (bonus[4] > 0) {
      if (item.university === "성신여자대학교") {
        if (sci1 || sci2) cases[4] = Math.max(sci1 ? scores.inquiry1 : 0, sci2 ? scores.inquiry2 : 0) * bonus[4];
      } else if (item.university === "건국대학교_글로컬") {
        if (sci1 || sci2) {
          if (sci1 && (!sci2 || scores.inquiry1 > scores.inquiry2)) scores.inquiry1 *= 1 + bonus[4];
          else scores.inquiry2 *= 1 + bonus[4];
        }
      } else if (item.university === "을지대학교") {
        if (sci1 && !sci2) scores.inquiry1 *= 1 + bonus[4];
        if (sci2 && !sci1) scores.inquiry2 *= 1 + bonus[4];
      } else if (item.university === "경희대학교") {
        if (sci1) scores.inquiry1 += bonus[4];
        if (sci2) scores.inquiry2 += bonus[4];
      } else {
        if (sci1) scores.inquiry1 *= 1 + bonus[4];
        if (sci2) scores.inquiry2 *= 1 + bonus[4];
      }
    }
    if (bonus[5] > 0) {
      const add = item.university === "경희대학교" ? bonus[5] : null;
      if (item.university === "서울시립대학교") {
        if (!sci1 && !sci2) { scores.inquiry1 *= 1 + bonus[5]; scores.inquiry2 *= 1 + bonus[5]; }
      } else {
        if (!sci1) scores.inquiry1 = add == null ? scores.inquiry1 * (1 + bonus[5]) : scores.inquiry1 + add;
        if (!sci2) scores.inquiry2 = add == null ? scores.inquiry2 * (1 + bonus[5]) : scores.inquiry2 + add;
      }
    }
    if (bonus[6] > 0 && sci1 && sci2) { scores.inquiry1 *= 1 + bonus[6]; scores.inquiry2 *= 1 + bonus[6]; }
    const level1 = scores.inquiry1Subject.endsWith("Ⅰ");
    const level2 = scores.inquiry2Subject.endsWith("Ⅰ");
    const adv1 = scores.inquiry1Subject.endsWith("Ⅱ");
    const adv2 = scores.inquiry2Subject.endsWith("Ⅱ");
    if (bonus[7] > 0) {
      if (item.university === "경상국립대학교" && level1 && level2) { scores.inquiry1 *= 1 + bonus[7]; scores.inquiry2 *= 1 + bonus[7]; }
      if (item.university === "국립부경대학교") cases[7] = (level1 ? scores.inquiry1 * (1 + bonus[7]) : 0) + (level2 ? scores.inquiry2 * (1 + bonus[7]) : 0);
    }
    if (bonus[8] > 0) {
      if (item.university === "서강대학교") cases[8] = (adv1 ? bonus[8] : 0) + (adv2 ? bonus[8] : 0);
      else if (item.university === "국립부경대학교") cases[8] = (adv1 ? scores.inquiry1 * (1 + bonus[8]) : 0) + (adv2 ? scores.inquiry2 * (1 + bonus[8]) : 0);
      else { if (adv1) scores.inquiry1 *= 1 + bonus[8]; if (adv2) scores.inquiry2 *= 1 + bonus[8]; }
    }
    if (bonus[9] > 0 && ((level1 && adv2) || (adv1 && level2))) {
      if (item.university === "서울대학교") cases[9] = bonus[9];
      if (item.university === "경상국립대학교") { scores.inquiry1 *= 1 + bonus[9]; scores.inquiry2 *= 1 + bonus[9]; }
    }
    if (bonus[10] > 0 && adv1 && adv2) {
      if (item.university === "서울대학교") cases[10] = bonus[10];
      if (item.university === "경상국립대학교") { scores.inquiry1 *= 1 + bonus[10]; scores.inquiry2 *= 1 + bonus[10]; }
    }
    if (bonus[11] > 0) {
      if (/물리학/.test(scores.inquiry1Subject)) scores.inquiry1 *= 1 + bonus[11];
      if (/물리학/.test(scores.inquiry2Subject)) scores.inquiry2 *= 1 + bonus[11];
    }
    if (bonus[13] > 0 && sci1 !== sci2) {
      if (sci1) scores.inquiry1 *= 1 + bonus[13];
      if (sci2) scores.inquiry2 *= 1 + bonus[13];
    }
    return cases;
  }

  function inquiryComposite(item, scores) {
    if (item.university === "가천대학교") return Math.max(scores.inquiry1, scores.inquiry2) * 2;
    if (["을지대학교", "성신여자대학교"].includes(item.university)) return avg(scores.inquiry1, scores.inquiry2);
    if (item.university === "삼육대학교" && !/2026/.test(scores.targetExamYear)) {
      if (["아트앤디자인", "음악학과"].includes(item.major)) return Math.max(scores.history, scores.inquiry1, scores.inquiry2);
      return scores.history > Math.min(scores.inquiry1, scores.inquiry2)
        ? avg(scores.history, Math.max(scores.inquiry1, scores.inquiry2))
        : avg(scores.inquiry1, scores.inquiry2);
    }
    return Math.max(scores.inquiry1, scores.inquiry2);
  }

  function buildRatios(item, scores) {
    const primary = [], alternate = [];
    for (const raw of item.ratioColumns || []) {
      const [a, b] = splitRatio(raw);
      primary.push(a); alternate.push(b);
    }
    while (primary.length < 27) { primary.push(0); alternate.push(0); }
    const explore = inquiryComposite(item, scores);
    const inquirySum = scores.inquiry1 + scores.inquiry2;
    const specialExplore = ["홍익대학교", "홍익대학교_세종", "이화여자대학교"].includes(item.university)
      ? inquirySum : item.university === "서경대학교" ? avg(scores.inquiry1, scores.inquiry2) : Math.max(scores.inquiry1, scores.inquiry2);
    const four = [scores.korean, scores.math, scores.english, explore];
    const values = [
      scores.korean, scores.math, scores.english, Math.max(scores.inquiry1, scores.inquiry2), inquirySum, scores.history,
      Math.max(scores.korean, scores.math), Math.min(scores.korean, scores.math),
      Math.max(scores.korean, scores.math, specialExplore), largest([scores.korean, scores.math, specialExplore], 2), Math.min(scores.korean, scores.math, specialExplore),
      largest(four, 1), largest(four, 2), largest(four, 3), largest(four, 4),
      item.university === "을지대학교" ? Math.max(scores.math, avg(scores.inquiry1, scores.inquiry2))
        : ["건국대학교", "국민대학교"].includes(item.university) ? Math.max(scores.math, inquirySum)
        : Math.max(scores.math, specialExplore),
      item.university === "을지대학교" ? Math.max(scores.korean, avg(scores.inquiry1, scores.inquiry2)) : Math.max(scores.korean, Math.max(scores.inquiry1, scores.inquiry2)),
      Math.min(scores.korean, Math.max(scores.inquiry1, scores.inquiry2)),
      largest([scores.korean, scores.english, Math.max(scores.inquiry1, scores.inquiry2)], 1), largest([scores.korean, scores.english, Math.max(scores.inquiry1, scores.inquiry2)], 2),
      largest([scores.math, scores.english, Math.max(scores.inquiry1, scores.inquiry2)], 1), largest([scores.math, scores.english, Math.max(scores.inquiry1, scores.inquiry2)], 2),
      Math.max(scores.math, scores.english), Math.max(scores.korean, scores.english),
      largest([scores.korean, scores.english, scores.math], 1), largest([scores.korean, scores.english, scores.math], 2), largest([scores.korean, scores.english, scores.math], 3),
    ];
    return { primary, alternate, values };
  }

  const dot = (weights, values, divisor = 1) => weights.reduce((sum, weight, index) => sum + n(weight) * n(values[index]) / divisor, 0);
  const dotExploreAverage = (weights, values, divisor = 100) => weights.reduce((sum, weight, index) => sum + n(weight) * n(values[index]) / divisor / (index === 4 ? 2 : 1), 0);

  function universityScore(item, scores, ratios, bonus, inputs) {
    const w = ratios.primary, wa = ratios.alternate, v = ratios.values;
    const max = n(item.fullScore, 100), year = Number(String(scores.targetExamYear).slice(0, 4));
    const uni = item.university;
    let total = 0;
    const normal100 = () => dot(w, v, 100);
    const average100 = () => dotExploreAverage(w, v, 100);

    if (uni === "성균관대학교") total = Math.max(dot(w, v), dot(wa, v));
    else if (uni === "가천대학교") total = average100() + scores.history;
    else if (uni === "가톨릭대학교") {
      total = normal100();
      if (year === 2024 && ["의예과", "약학과", "간호학과(인문)", "간호학과(자연)"].includes(item.major)) total = total * 5 + scores.english + scores.history;
      else if (year >= 2025 && item.major === "의예과") total = (total * 5 + scores.english + scores.history) * 0.95;
      else if (year >= 2025 && ["간호학과", "약학과"].includes(item.major)) total = total * 5 + scores.english + scores.history;
      else total = total * 10 + scores.history;
    } else if (uni === "강원대학교_춘천") total = average100() * max / 100 + scores.history;
    else if (uni === "건국대학교") total = Math.max(dot(w, v, 100), dot(wa, v, 100)) * max / 200 + scores.history;
    else if (uni === "경기대학교") total = Math.max(dot(w, v, 100), dot(wa, v, 100)) * max / 100 + scores.history;
    else if (uni === "경북대학교") {
      let divisor = 0;
      total = w.reduce((sum, weight, i) => {
        divisor += n(weight);
        if (!positive(weight)) return sum;
        if (i === 2) return sum + v[i] * (weight === 200 ? 2 : 1);
        if (i === 3) return sum + weight * v[i] / 100;
        return sum + weight * v[i] / 200;
      }, 0) * max / divisor + scores.history;
    }
    else if (uni === "경희대학교") {
      total = w[0] === 50 || w[0] === 60 ? w.reduce((s, weight, i) => s + weight * (i === 3 ? v[i] + 100 : v[i]) / 100, 0) : normal100();
      total = year < 2026 ? (total + scores.history) / 2 * max / 100 : total / 2 * max / 100 + scores.english + scores.history;
    } else if (uni === "고려대학교") total = dot(w, v, 200) / w.reduce((s, x) => s + n(x), 0) * max + scores.history + scores.english;
    else if (uni === "고려대학교_세종") {
      const inquiryMax = year === 2024 ? 71.75 : year === 2025 ? 70.12 : 70.11;
      let earned = 0, possible = 0;
      for (let i = 0; i < w.length; i += 1) {
        if (!positive(w[i])) continue;
        earned += v[i] * w[i] / 100;
        possible += (i === 4 ? inquiryMax * w[i] * 2 : n(scores.maxima[i]) * w[i]) / 100;
      }
      total = earned / possible * max;
    }
    else if (uni === "광운대학교") total = normal100() * max / 200 + scores.history;
    else if (uni === "국립목포해양대학교") total = average100() * 10;
    else if (uni === "국립부경대학교") total = dot(w, v, 200) * 10 + bonus[2] + bonus[3] + bonus[7] + bonus[8] + scores.history;
    else if (uni === "국민대학교") total = w.reduce((s, weight, i) => s + weight * v[i] * (i === 2 ? 2 : 1), 0) / 200 * max / w.reduce((s, x) => s + n(x), 0) + scores.history;
    else if (uni === "단국대학교_죽전" || uni === "단국대학교_천안") {
      const medical = ["의예과", "치의예과", "약학과"].includes(item.major);
      total = w.reduce((sum, weight, i) => {
        if (!positive(weight)) return sum;
        if (medical && i === 0) return sum + weight * v[i] / Math.max(scores.maxima[0], 1);
        if (medical && i === 1) return sum + weight * v[i] / Math.max(scores.maxima[1], 1);
        return sum + weight * v[i] / 100 / (i === 4 ? 2 : 1);
      }, 0) * max / 100 + scores.history;
    } else if (["덕성여자대학교", "동덕여자대학교", "서경대학교", "총신대학교"].includes(uni)) total = average100() * max / 100 + scores.history;
    else if (uni === "동국대학교" || uni === "동국대학교_경주") total = (uni === "동국대학교_경주" ? average100() / 100 : normal100() / 200) * max + scores.history;
    else if (uni === "명지대학교") total = normal100() * max / 100 + scores.history;
    else if (uni === "부산대학교") total = w.reduce((s, weight, i) => i === 2 ? s : s + weight * v[i] / 200, 0) + scores.english + scores.history;
    else if (uni === "삼육대학교") total = average100() * max / 100;
    else if (uni === "상명대학교" || uni === "상명대학교_천안") total = Math.min(max, normal100() * max / 100 + scores.history);
    else if (uni === "서강대학교") total = Math.max(dot(w, v), dot(wa, v)) + w[4] * bonus[8] + scores.english + scores.history;
    else if (uni === "서울교육대학교") total = dot(w, v, 160);
    else if (uni === "서울대학교") {
      total = dot(w, v) + scores.english + scores.history + bonus[9] + bonus[10] + scores.secondLanguage;
      if (year <= 2025) {
        const grade = Math.max(1, Math.min(9, Math.trunc(n(inputs.math.grade, 9))));
        if (["동양화과", "서양화과", "조소과", "공예과", ...(year === 2024 ? ["디자인과"] : [])].includes(item.major)) total += [0, -0.5, -2, -4, -6, -8, -10, -12, -14][grade - 1];
        if (item.major === "성악과") total += [0, 0, 0, 0, -0.4, -0.8, -1.2, -1.6, -2][grade - 1];
        if (item.major === "작곡과") total += [0, -0.5, -1, -1.5, -2, -2.5, -3, -3.5, -4][grade - 1];
      }
    } else if (["서울여자대학교", "서울한영대학교", "성공회대학교"].includes(uni)) total = normal100() * max / 100 + (uni === "서울한영대학교" ? 0 : scores.history);
    else if (uni === "성신여자대학교") total = average100() * max / 100 + bonus[4] + scores.history;
    else if (uni === "세종대학교") total = normal100() * max / 200 + scores.history;
    else if (uni === "연세대학교") total = w.reduce((s, weight, i) => s + weight * v[i] / (i === 2 ? 100 : 200), 0) * max / w.reduce((s, x) => s + n(x), 0) + scores.history;
    else if (uni === "연세대학교_미래") total = w.reduce((s, weight, i) => s + weight * v[i] * (i === 3 ? 2 : 1) / (i === 2 ? 100 : 200), 0) * max / w.reduce((s, x) => s + n(x), 0) + scores.history;
    else if (uni === "을지대학교") {
      total = average100();
      if (item.track === "일반전형Ⅰ") total = 90 + total * 8.1;
      if (item.track === "일반전형Ⅱ") total *= 10;
      total += scores.history;
    } else if (uni === "인천대학교") total = dotExploreAverage(w, v, 10) * max / 1000 + scores.history;
    else if (uni === "전북대학교") total = w.reduce((s, weight, i) => s + weight * v[i] / (i === 4 ? 2 : 1), 0) + scores.english + scores.history;
    else if (uni === "중앙대학교") total = normal100() * max / 200 + scores.history + scores.english;
    else if (uni === "충남대학교") total = dot(w, v) / 200 + scores.english + scores.history;
    else if (uni === "한국교원대학교") {
      total = w.reduce((s, weight, i) => s + weight * v[i] / (i === 4 ? 2 : 1), 0);
      if (item.track === "예술체육실기") total *= 260 / 300;
    } else if (uni === "한남대학교") total = Math.min(max, v.reduce((s, x) => s + n(x), 0) * max / 300 + scores.history);
    else if (uni === "한성대학교") total = w.reduce((s, weight, i) => s + weight * v[i] * (i === 2 && v[i] === 100 && weight === 20 ? 2 : 1) / 100, 0) * max / 100 + scores.history;
    else if (uni === "홍익대학교" || uni === "홍익대학교_세종") total = normal100() + scores.history;
    else total = normalizedUniversityScore(item, scores, ratios, bonus);
    return excelRound(total, uni === "상명대학교" ? 3 : 2);
  }

  function normalizedUniversityScore(item, scores, ratios, bonus) {
    const w = ratios.primary, wa = ratios.alternate, v = ratios.values;
    const max = n(item.fullScore, 100), uni = item.university, year = Number(String(scores.targetExamYear).slice(0, 4));
    const m = scores.maxima.map((value) => Math.max(n(value), 1));
    let total = 0;
    if (uni === "서울시립대학교") total = w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[2] * scores.english / m[2] + w[4] / 2 * (scores.inquiry1 / m[3] + scores.inquiry2 / m[4]) + scores.history;
    else if (uni === "이화여자대학교") {
      const calc = (weights) => weights[0] * scores.korean / m[0] / 100 + weights[1] * scores.math / m[1] / 100 + weights[2] * scores.english / m[2] / 100 + weights[3] * Math.max(scores.inquiry1, scores.inquiry2) / Math.max(m[3], m[4]) / 100 + weights[4] * (scores.inquiry1 / m[3] + scores.inquiry2 / m[4]) / 200;
      total = Math.max(calc(w) * max + scores.history, calc(wa) * max + scores.historyOther);
    } else if (uni === "인하대학교") {
      const calc = (weights) => weights[0] * scores.korean / m[0] + weights[1] * scores.math / m[1] + weights[2] * scores.english / m[2] + weights[4] * (scores.inquiry1 + scores.inquiry2) / 2 / Math.max(m[3], m[4]) + weights[5] * scores.history / 50;
      total = Math.max(calc(w), calc(wa)) * max / 100 + (year >= 2025 ? scores.history : 0);
    } else if (uni === "전남대학교") total = (w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[3] * Math.max(scores.inquiry1 / m[3], scores.inquiry2 / m[4]) + w[4] * avg(scores.inquiry1 / m[3], scores.inquiry2 / m[4])) / 100 * (max - m[2]) + scores.english + scores.history;
    else if (uni === "충북대학교") total = (w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[2] * scores.english / m[2] + w[4] * (scores.inquiry1 + scores.inquiry2) / (m[3] + m[4])) / 100 * 200 + 800 + scores.history;
    else if (uni === "한국외국어대학교" || uni === "한국외국어대학교_글로벌캠퍼스") {
      const inquiryMax = year === 2024 ? 69.35 : year === 2025 ? 70 : 70.12;
      total = (w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[4] * (scores.inquiry1 + scores.inquiry2) / (inquiryMax * 2)) / 100 * 700 + scores.english + scores.history;
    } else if (uni === "숙명여자대학교") {
      const inquiryMax = year === 2024 ? 69.35 : year === 2025 ? 70 : 70.11;
      total = (w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[2] * scores.english / m[2] + w[3] * Math.max(scores.inquiry1, scores.inquiry2) / inquiryMax + w[4] * (scores.inquiry1 + scores.inquiry2) / (inquiryMax * 2)) / 100 * max + scores.history;
    } else if (uni === "숭실대학교") {
      const inquiryMax = year === 2024 ? 71.75 : year === 2025 ? 70 : 70.12;
      total = w.reduce((sum, weight, i) => sum + weight * v[i] / (i < 3 ? m[i] : inquiryMax * 2) * 10, 0) + scores.history;
    } else if (uni === "아주대학교") total = (w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[2] * scores.english / m[2] + w[4] * (scores.inquiry1 + scores.inquiry2) / (m[3] + m[4])) / 100 * 1000 + scores.history;
    else if (uni === "한양대학교" || uni === "한양대학교_에리카") {
      const inquiryMax = uni === "한양대학교" ? (year === 2024 ? 68.85 : year === 2025 ? 70 : 69.44) : (year === 2024 ? 68.85 : year === 2025 ? 70 : 70.12);
      total = w[0] * scores.korean / m[0] * 10 + w[1] * scores.math / m[1] * 10 + (w[2] > 0 ? scores.english : 0) + w[3] * Math.max(scores.inquiry1, scores.inquiry2) / inquiryMax * 10 + w[4] * (scores.inquiry1 + scores.inquiry2) / (inquiryMax * 2) * 10;
      total = uni === "한양대학교" ? total * max / 1000 + scores.history : total + w[15] * Math.max(scores.math / m[1], Math.max(scores.inquiry1, scores.inquiry2) / inquiryMax) * 10 + scores.history;
    } else if (uni === "서울과학기술대학교") total = (w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + w[2] * scores.english / m[2] + w[4] * avg(scores.inquiry1 / m[3], scores.inquiry2 / m[4]) + w[15] * Math.max(scores.math / m[1], avg(scores.inquiry1 / m[3], scores.inquiry2 / m[4]))) / 100 * 1000 + scores.history;
    else if (uni === "경상국립대학교") total = w[0] * scores.korean / m[0] + w[1] * scores.math / m[1] + (w[2] > 0 ? scores.english * max / 1000 : 0) + w[4] * avg(scores.inquiry1 / m[3], scores.inquiry2 / m[4]) + scores.math * (bonus[2] + bonus[3]) + v[4] * (bonus[7] + bonus[9] + bonus[10]);
    else if (uni === "건국대학교_글로컬") {
      total = w[0] * scores.korean / m[0] / 100 + w[1] * scores.math / m[1] / 100 + w[2] * scores.english / m[2] / 100 + w[4] * avg(scores.inquiry1 / m[3], scores.inquiry2 / m[4]) / 100;
      for (let i = 0; i < w.length; i += 1) if (![0, 1, 2, 4].includes(i)) total += w[i] * v[i] / 100 * (year === 2026 ? 10 : 1);
      if (item.major.includes("의예과")) total *= max;
      total = Math.min(max, total + scores.history);
    }
    else total = Math.max(dot(w, v, 100), dot(wa, v, 100)) * max / 100 + scores.history;
    return total;
  }

  function chanceLabel(myScore, cut50, cut70) {
    const my = excelRound(myScore, 5), c50 = positive(cut50) ? n(cut50) : 0, c70 = positive(cut70) ? n(cut70) : 0;
    const gap = 1.65;
    if (c50 && c70) {
      const difference = c50 - c70;
      if (c50 === c70) {
        if (my >= c50 + 2 * gap) return "80% 이상";
        if (my >= c50 + gap) return "70% 이상";
        if (my >= c50) return "60% 이상";
        if (my >= c50 - gap) return "50% 이상";
        if (my >= c50 - 2 * gap) return "40% 이상";
        if (my >= c50 - 3 * gap) return "30% 이상";
        return "30% 미만";
      }
      if (my >= c50) return my >= c50 + difference ? "90% 이상" : my >= c50 + difference / 2 ? "80% 이상" : "70% 이상";
      if (my >= c70) return `${excelRound(60 + 20 * (my - c70) / difference, -1)}% 이상`;
      if (my >= c70 - difference / 2) return "50% 이상";
      if (my >= c70 - difference) return "40% 이상";
      if (my >= c70 - difference * 1.5) return "30% 이상";
      return "30% 미만";
    }
    const cut = c50 || c70;
    if (!cut) return "정보없음";
    const offsets = c50 ? [1, 0, -1, -2, -3, -4] : [3, 2, 1, 0, -1, -2, -3];
    const labels = c50 ? ["80% 이상", "70% 이상", "60% 이상", "50% 이상", "40% 이상", "30% 이상"] : ["90% 이상", "80% 이상", "70% 이상", "60% 이상", "50% 이상", "40% 이상", "30% 이상"];
    for (let i = 0; i < offsets.length; i += 1) if (my >= cut + offsets[i] * gap) return labels[i];
    return "30% 미만";
  }

  function calculate(item, inputs) {
    const scores = processScores(item, inputs);
    const bonus = applyBonuses(item, scores, inputs);
    const ratios = buildRatios(item, scores);
    const myScore = universityScore(item, scores, ratios, bonus, inputs);
    const label = chanceLabel(myScore, item.cut50, item.cut70);
    return { myScore, chanceLabel: label, chance: n(label.match(/\d+/)?.[0], 0), scores };
  }

  return { calculate, chanceLabel };
}

