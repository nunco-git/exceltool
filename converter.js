/* 세출합본예산서 변환 로직 (엑셀 매크로 Module1.bas 를 그대로 옮긴 것)
 * 입력은 "격자(grid)" = 행 배열의 배열. 칸 값은 문자열/숫자/빈칸(null·undefined·'') 중 하나.
 */
(function (root) {
  'use strict';
  var OUT_COLS = 32;
  var CIRCLE = '○'; // ○

  // ---------- VBA 흉내 도우미 ----------
  function isEmpty(x) { return x === null || x === undefined || x === ''; }
  function rtrim(s) { return s.replace(/ +$/, ''); }
  function ltrim(s) { return s.replace(/^ +/, ''); }
  function trim(s) { return s.replace(/^ +| +$/g, ''); }
  function isNum(s) {
    s = trim(s);
    return /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s) || /^\([0-9.]+\)$/.test(s);
  }
  function toNum(s) {
    s = trim(s);
    if (/^\(.*\)$/.test(s)) return -parseFloat(s.slice(1, -1));
    return parseFloat(s);
  }
  function joinWrap(x, sep) {
    if (isEmpty(x)) return '';
    var s = String(x).replace(/\r/g, '');
    var p = s.split('\n');
    var r = p[0];
    for (var i = 1; i < p.length; i++) {
      var t = trim(p[i]);
      if (t.length > 0) r = sep.length === 0 ? r + t : rtrim(r) + sep + t;
    }
    return r;
  }
  function isAmt(x) {
    if (isEmpty(x)) return false;
    if (typeof x === 'number') return true;
    var s = trim(String(x).replace(/[\r\n]/g, '')).replace(/,/g, '');
    return s.length > 0 && isNum(s);
  }
  function amt(x) {
    if (isEmpty(x)) return 0;
    if (typeof x === 'number') return x;
    var s = trim(String(x).replace(/\n/g, '').replace(/\r/g, '').replace(/,/g, ''));
    if (s.length > 0 && isNum(s)) return toNum(s);
    return 0;
  }
  function pad(n, w) { var s = String(Math.round(n)); while (s.length < w) s = '0' + s; return s; }
  function codeOf(x) {
    if (isEmpty(x)) return '';
    if (typeof x === 'number') return x >= 100 ? pad(x, 3) : pad(x, 2);
    var s = trim(joinWrap(x, ''));
    if (/^\d{1,2}$/.test(s)) s = pad(parseInt(s, 10), 2); // 붙여넣기 중 앞자리 0이 사라진 경우 복원
    return s;
  }
  function txtVal(x) {
    if (isEmpty(x)) return undefined;
    var s = typeof x === 'number' ? String(Math.round(x)) : String(x).replace(/\r/g, '').replace(/\n/g, '');
    return s.length > 0 ? s : undefined;
  }
  function S(x) { return x === undefined || x === null ? '' : String(x); }
  function normTxt(s) { return s.replace(/ /g, '').replace(/ /g, ''); }
  function fmt(n) { return Math.round(n).toLocaleString('en-US'); }

  function lastRowOf(grid) {
    for (var r = grid.length - 1; r >= 0; r--) {
      var row = grid[r];
      if (row) for (var c = 0; c < row.length; c++) if (!isEmpty(row[c])) return r + 1;
    }
    return 0;
  }
  function lastColOf(grid) {
    var m = 0;
    for (var r = 0; r < grid.length; r++) {
      var row = grid[r];
      if (row) for (var c = row.length - 1; c >= m; c--) if (!isEmpty(row[c])) { m = c + 1; break; }
    }
    return m;
  }
  function cell(grid, r, c) { var row = grid[r - 1]; return row ? row[c - 1] : undefined; }
  function hCol(grid, nm, ncol) {
    for (var c = 1; c <= ncol; c++) {
      var x = cell(grid, 1, c);
      if (!isEmpty(x) && trim(String(x)) === nm) return c;
    }
    return 0;
  }
  function splitPolicy(s) {
    var res = { nm: s, fld: '', sec: '' };
    if (s.slice(-1) !== ')') return res;
    var p = s.lastIndexOf('(') + 1;
    if (p < 2) return res;
    var inner = s.substring(p, s.length - 1);
    var k = inner.indexOf('/') + 1;
    if (k === 0) return res;
    res.nm = trim(s.substring(0, p - 1));
    res.fld = trim(inner.substring(0, k - 1));
    res.sec = trim(inner.substring(k));
    return res;
  }
  function lookupDept(depts, nm) {
    for (var i = 0; i < depts.length; i++) {
      if (trim(S(depts[i].name)) === nm) {
        var d = depts[i];
        return { found: true, code: trim(S(d.code)), com: trim(S(d.com)), acc: trim(S(d.acc)), accNm: trim(S(d.accNm)) };
      }
    }
    return { found: false, code: '', com: '', acc: '', accNm: '' };
  }
  function opt(s) { return s.length > 0 ? s : undefined; }
  var DEFAULT_ACCTS = [{ name: '일반회계', code: '100' }, { name: '양성평등기금', code: '411' }];
  // 단위사업명 뒤에 붙은 (회계명) 분리: 예) "양성평등 기반조성(양성평등기금)"
  function splitUnitAcct(unit, accts) {
    var m = String(unit).match(/^(.*?)\s*\(([^()]*)\)\s*$/);
    if (m) {
      for (var i = 0; i < accts.length; i++) {
        if (trim(S(accts[i].name)) === trim(m[2])) return { unit: trim(m[1]), accNm: trim(S(accts[i].name)), acc: trim(S(accts[i].code)) };
      }
    }
    return null;
  }


  // ---------- 줄바꿈으로 잘린 글자 이어붙이기 ----------
  // 보고서는 긴 글자를 40칸쯤에서 꺾어 다음 줄에 이어 쓴다. 꺾은 자리가 띄어쓰기였는지 단어 중간이었는지는
  // 원본에 남아 있지 않으므로, 같은 파일 안의 글자 연결 빈도와 줄 길이로 추정한다.
  function bw(s) { var w = 0; for (var i = 0; i < s.length; i++) w += s.charCodeAt(i) > 255 ? 2 : 1; return w; }
  function chars(s) { return Array.from(s); }
  function makeJoiner(grid, c1, c2) {
    var pin = {}, psp = {};
    function learn(line) {
      var t = trim(line).replace(/^[○\-ㆍ]+/, '');
      var ws = t.split(/\s+/);
      var i, cs;
      for (i = 0; i < ws.length; i++) { cs = chars(ws[i]); for (var k = 0; k < cs.length - 1; k++) { var key = cs[k] + cs[k + 1]; pin[key] = (pin[key] || 0) + 1; } }
      for (i = 0; i < ws.length - 1; i++) {
        var a = chars(ws[i]), b = chars(ws[i + 1]);
        if (a.length && b.length) { var kk = a[a.length - 1] + b[0]; psp[kk] = (psp[kk] || 0) + 1; }
      }
    }
    for (var r = 0; r < grid.length; r++) {
      var row = grid[r]; if (!row) continue;
      for (var c = c1 - 1; c <= c2 - 1 && c < row.length; c++) {
        var x = row[c];
        if (typeof x === 'string' && x.length > 0) x.replace(/\r/g, '').split('\n').forEach(function (p) { if (trim(p).length) learn(p); });
      }
    }
    return function (x) {
      if (isEmpty(x)) return '';
      var p = String(x).replace(/\r/g, '').split('\n');
      var r = rtrim(p[0]);
      for (var i = 1; i < p.length; i++) {
        var t = trim(p[i]);
        if (t.length === 0) continue;
        var a = rtrim(p[i - 1]);
        var ca = chars(trim(a)), cb = chars(t);
        var sep = ' ';
        if (ca.length && cb.length) {
          var key = ca[ca.length - 1] + cb[0];
          var sc = Math.log(((pin[key] || 0) + 0.5) / ((psp[key] || 0) + 0.5)) + (bw(a) >= 39 ? 0.6 : -0.8);
          if (sc > 0) sep = '';
        }
        r = r + sep + t;
      }
      return r;
    };
  }

  // ---------- 2026 자료(시트3) 읽기 ----------
  function loadOld(grid3, set) {
    var o = { m: 0, why: '', hasCtx: false };
    if (!grid3) { o.why = '2026 전체자료가 입력되지 않아 줄별 기정액을 확인하지 못했습니다.'; return o; }
    var lastRow = lastRowOf(grid3), lastCol = lastColOf(grid3);
    if (lastRow > 0 && lastCol < 30 && isDetailGrid(grid3)) return loadOldDetail(grid3, set || {});
    if (lastRow === 0) { o.why = '2026 전체자료가 비어 있어 줄별 기정액을 확인하지 못했습니다.'; return o; }
    if (lastRow < 3 || lastCol < 30) { o.why = '2026 전체자료 형식이 맞지 않아 줄별 기정액을 확인하지 못했습니다.'; return o; }
    var cD = hCol(grid3, '부서명', lastCol), cP = hCol(grid3, '세부사업명', lastCol), cT = hCol(grid3, '통계목', lastCol);
    var cN = hCol(grid3, '산출근거명', lastCol), cX = hCol(grid3, '산출근거식', lastCol), cA = hCol(grid3, '본예산', lastCol);
    var en = ['분야명', '부문명', '정책사업명', '단위사업명', '편성목', '편성목명', '통계목명', '회계코드', '회계명'];
    var e = [];
    o.hasCtx = true;
    for (var q = 0; q < 9; q++) { e[q] = hCol(grid3, en[q], lastCol); if (e[q] === 0) o.hasCtx = false; }
    if (cD * cP * cT * cN * cX * cA === 0) { o.why = "2026 전체자료에서 '본예산' 등 제목 열을 찾지 못해 줄별 기정액을 확인하지 못했습니다."; o.hasCtx = false; return o; }
    var dp = [], pj = [], tg = [], nm = [], raw = [], fx = [], ak = [];
    var mFld = [], mSec = [], mPol = [], mUnit = [], mPy = [], mPyN = [], mTgN = [], mAcc = [], mAccN = [];
    var m = 0;
    for (var i = 3; i <= lastRow; i++) {
      var a = amt(cell(grid3, i, cA));
      if (a !== 0) {
        m++;
        dp[m] = trim(S(txtVal(cell(grid3, i, cD))));
        pj[m] = S(txtVal(cell(grid3, i, cP)));
        tg[m] = S(txtVal(cell(grid3, i, cT)));
        raw[m] = S(txtVal(cell(grid3, i, cN)));
        nm[m] = normTxt(raw[m]);
        fx[m] = S(txtVal(cell(grid3, i, cX)));
        ak[m] = a;
        if (o.hasCtx) {
          mFld[m] = S(txtVal(cell(grid3, i, e[0]))); mSec[m] = S(txtVal(cell(grid3, i, e[1]))); mPol[m] = S(txtVal(cell(grid3, i, e[2])));
          mUnit[m] = S(txtVal(cell(grid3, i, e[3]))); mPy[m] = S(txtVal(cell(grid3, i, e[4]))); mPyN[m] = S(txtVal(cell(grid3, i, e[5])));
          mTgN[m] = S(txtVal(cell(grid3, i, e[6]))); mAcc[m] = S(txtVal(cell(grid3, i, e[7]))); mAccN[m] = S(txtVal(cell(grid3, i, e[8])));
        }
      }
    }
    if (m === 0) { o.why = '2026 전체자료에 본예산 금액이 있는 줄이 없습니다.'; o.hasCtx = false; return o; }
    o.m = m; o.dp = dp; o.pj = pj; o.tg = tg; o.nm = nm; o.raw = raw; o.fx = fx; o.ak = ak;
    o.fld = mFld; o.sec = mSec; o.pol = mPol; o.unit = mUnit; o.py = mPy; o.pyN = mPyN; o.tgN = mTgN; o.acc = mAcc; o.accN = mAccN;
    return o;
  }

  // 보고서형(15100) 파일에서 본예산 줄을 읽어 loadOld 와 같은 모양으로 돌려준다
  function loadOldDetail(grid3, set) {
    var accts = set.accts || DEFAULT_ACCTS;
    var P = parseDetail(grid3, set.depts || []);
    var o = { m: 0, why: '', hasCtx: true };
    var dp = [], pj = [], tg = [], nm = [], raw = [], fx = [], ak = [];
    var mFld = [], mSec = [], mPol = [], mUnit = [], mPy = [], mPyN = [], mTgN = [], mAcc = [], mAccN = [];
    var m = 0;
    P.lines.forEach(function (l) {
      if (l.block !== '본예산' || l.parent || l.b === 0) return;
      m++;
      var sp = splitPolicy(l.pol), ua = splitUnitAcct(l.unit, accts);
      dp[m] = trim(l.dept); pj[m] = l.proj; tg[m] = l.pcode + '-' + l.tcode; raw[m] = l.text; nm[m] = normTxt(l.text);
      fx[m] = l.formula; ak[m] = l.b;
      mFld[m] = sp.fld; mSec[m] = sp.sec; mPol[m] = sp.nm; mUnit[m] = ua ? ua.unit : l.unit; mPy[m] = l.pcode; mPyN[m] = l.pname;
      mTgN[m] = l.tname; mAcc[m] = ua ? ua.acc : l.ctx.acc; mAccN[m] = ua ? ua.accNm : l.ctx.accNm;
    });
    if (m === 0) { o.why = '2026 전체자료에 본예산 금액이 있는 줄이 없습니다.'; o.hasCtx = false; return o; }
    o.m = m; o.dp = dp; o.pj = pj; o.tg = tg; o.nm = nm; o.raw = raw; o.fx = fx; o.ak = ak;
    o.fld = mFld; o.sec = mSec; o.pol = mPol; o.unit = mUnit; o.py = mPy; o.pyN = mPyN; o.tgN = mTgN; o.acc = mAcc; o.accN = mAccN;
    return o;
  }

  function sumOld(O, o, d, p) {
    var t = 0;
    for (var j = 1; j <= O.m; j++) if (O.pj[j] === o && O.dp[j] === d && O.tg[j] === p) t += O.ak[j];
    return t;
  }

  // ---------- 시트1(예산요구서) -> 시트2 ----------
  function convertReport(grid1, grid3, set) {
    set = set || {};
    var depts = set.depts || [], accts = set.accts || DEFAULT_ACCTS, uAcc = '', uAccNm = '';
    var cfmDate = trim(S(set.cfmDate));
    var gubunSet = trim(S(set.gubun));
    var addExtra = set.addExtra !== false;
    var lastRow = lastRowOf(grid1);
    if (lastRow === 0) return { error: '예산요구서(시트1)가 비어 있습니다.' };
    var v = function (i, c) { return cell(grid1, i, c); };
    var joiner = makeJoiner(grid1, 2, 7);
    var gcol = 10, lineMode = false, lineK = [];
    for (var hr = 1; hr <= 12; hr++) for (var hc = 9; hc <= 12; hc++) {
      var hv = v(hr, hc);
      if (!isEmpty(hv) && String(hv).replace(/\s/g, '') === '기정액') gcol = hc;
    }
    var i, k, c, g, j, a, b, cc, d;

    var titleSrc = '';
    for (i = 1; i <= 8; i++) {
      if (i > lastRow) break;
      a = joinWrap(v(i, 1), ' ');
      if (a.indexOf('년') >= 0 && (a.indexOf('예산') >= 0 || a.indexOf('추경') >= 0)) { titleSrc = trim(a); break; }
    }
    var pre = titleSrc;
    k = titleSrc.indexOf('일반회계') + 1;
    if (k > 1) pre = trim(titleSrc.substring(0, k - 1));
    pre = pre.split('년도 ').join('년 ');
    var titleOut = pre.length > 0 ? pre + ' 세출합본예산서(일반,특별)' : '세출합본예산서(일반,특별)';
    var gubun = gubunSet;
    if (gubun.length === 0) {
      if (pre.indexOf('본예산') >= 0) gubun = '본예산';
      else if (pre.indexOf('추경') >= 0) {
        gubun = '추경';
        var q = pre.indexOf('회') + 1;
        if (q > 1 && /\d/.test(pre.charAt(q - 2))) gubun = '추경' + pre.charAt(q - 2) + '회';
      }
    }

    var mm = titleSrc.match(/(\d{4})\s*년(?:도)?\s*추경\s*(\d+)\s*회/);
    if (mm) {
      titleOut = mm[1] + '년 제' + mm[2] + '회 추경 세출합본예산서(일반,특별)';
      if (gubunSet.length === 0) gubun = '추경' + mm[2] + '회';
    }

    var out = [], gStart = [], gCnt = [], gBud = [], gU = [], gSum = [], gKO = [], gKP = [], gDp = [], gCtx = [];
    var n = 0, gi = 0;
    var dept = '', pol = '', fld = '', sec = '', unit = '', proj = '', pyeong = '', pname = '', tong = '', tname = '';
    var dCode = '', dCom = '', dAcc = '', dAccNm = '';
    var sumDept = 0, sumDeptJ = 0, sumLines = 0, missing = [];

    for (i = 1; i <= lastRow; i++) {
      a = trim(joinWrap(v(i, 1), ''));
      b = trim(joinWrap(v(i, 2), ''));
      cc = trim(joinWrap(v(i, 3), ''));
      d = trim(joinWrap(v(i, 4), ''));
      if (a.length > 0) {
        if (isAmt(v(i, 9)) && a.slice(-1) !== ':') {
          dept = a; pol = ''; fld = ''; sec = ''; unit = ''; proj = ''; uAcc = ''; uAccNm = '';
          pyeong = ''; pname = ''; tong = ''; tname = '';
          sumDept += amt(v(i, 9));
          sumDeptJ += amt(v(i, gcol));
          var dd = lookupDept(depts, dept);
          dCode = dd.code; dCom = dd.com; dAcc = dd.acc; dAccNm = dd.accNm;
          if (!dd.found && missing.indexOf(dept) < 0) missing.push(dept);
        }
      } else if (b.length > 0) {
        var sp = splitPolicy(b);
        pol = sp.nm; fld = sp.fld; sec = sp.sec;
      } else if (cc.length > 0) {
        var ua1 = splitUnitAcct(cc, accts);
        if (ua1) { unit = ua1.unit; uAcc = ua1.acc; uAccNm = ua1.accNm; } else { unit = cc; uAcc = ''; uAccNm = ''; }
      } else if (d.length > 0) {
        proj = d;
      } else if (trim(joinWrap(v(i, 6), '')).length > 0) {
        var h = trim(joinWrap(v(i, 8), ''));
        if (h.length > 0) {
          var code = codeOf(v(i, 6));
          if (code.length === 3) { pyeong = code; pname = h; }
          else {
            tong = pyeong + '-' + code; tname = h;
            gi++;
            gStart[gi] = n + 1; gCnt[gi] = 0; gSum[gi] = 0;
            gBud[gi] = amt(v(i, 9)); gU[gi] = amt(v(i, gcol));
            gKO[gi] = proj; gKP[gi] = tong; gDp[gi] = dept;
            var cx = [];
            cx[1] = opt(cfmDate); cx[2] = opt(uAcc || dAcc); cx[3] = opt(dCode); cx[4] = opt(fld); cx[5] = opt(sec); cx[6] = opt(dCom);
            cx[7] = pol; cx[8] = unit; cx[10] = pyeong; cx[11] = pname; cx[12] = opt(gubun); cx[13] = opt(uAccNm || dAccNm);
            cx[14] = dept; cx[15] = proj; cx[16] = tong; cx[17] = tname;
            gCtx[gi] = cx;
          }
        } else {
          var txt = joiner(v(i, 6));
          var st = ltrim(txt);
          txt = st.charAt(0) === CIRCLE ? st : '  ' + st;
          var t = amt(v(i, 9));
          var u = 0;
          if (gi > 0 && gCnt[gi] === 0) u = gU[gi];
          n++;
          if (isAmt(v(i, gcol))) lineMode = true;
          lineK[n] = amt(v(i, gcol));
          var o = new Array(OUT_COLS + 1);
          o[1] = opt(cfmDate); o[2] = opt(uAcc || dAcc); o[3] = opt(dCode); o[4] = opt(fld); o[5] = opt(sec); o[6] = opt(dCom);
          o[7] = pol; o[8] = unit; o[10] = pyeong; o[11] = pname; o[12] = opt(gubun); o[13] = opt(uAccNm || dAccNm);
          o[14] = dept; o[15] = proj; o[16] = tong; o[17] = tname; o[18] = txt;
          var tx2 = trim(joinWrap(v(i, 7), ' '));
          if (tx2.length > 0) o[19] = tx2;
          o[20] = t; o[21] = u; o[22] = t - u;
          sumLines += t;
          if (gi > 0) { gCnt[gi]++; gSum[gi] += t; }
          out[n] = o;
        }
      }
    }

    // --- 줄별 기정액 보정 ---
    var O = lineMode ? { m: 0, why: '', hasCtx: false } : loadOld(grid3, { depts: depts, accts: accts });
    var m = O.m;
    var fin = [], flg = [], nFin = 0, i2 = 1;
    var used = [], effP = [], okg = [], clO = [], clD = [], clP = [], nCl = 0;
    var nAlt = 0, nVer = 0, nUnv = 0, nUnvLines = 0, nExtra = 0, nLeft = 0, nWhole = 0, nBlk = 0;
    if (m > 0) {
      for (g = 1; g <= gi; g++) {
        effP[g] = gKP[g];
        if (gU[g] !== 0) {
          if (sumOld(O, gKO[g], gDp[g], gKP[g]) === gU[g]) {
            okg[g] = true; nCl++; clO[nCl] = gKO[g]; clD[nCl] = gDp[g]; clP[nCl] = gKP[g];
          }
        }
      }
      for (g = 1; g <= gi; g++) {
        if (gU[g] !== 0 && !okg[g]) {
          var nCand = 0, candP = '';
          for (j = 1; j <= m; j++) {
            if (O.pj[j] === gKO[g] && O.dp[j] === gDp[g] && O.tg[j].substring(0, 3) === gKP[g].substring(0, 3) && O.tg[j] !== gKP[g]) {
              var isFirst = true, jj;
              for (jj = 1; jj <= j - 1; jj++) if (O.pj[jj] === O.pj[j] && O.dp[jj] === O.dp[j] && O.tg[jj] === O.tg[j]) { isFirst = false; break; }
              var claimed = false;
              for (jj = 1; jj <= nCl; jj++) if (clO[jj] === O.pj[j] && clD[jj] === O.dp[j] && clP[jj] === O.tg[j]) { claimed = true; break; }
              if (isFirst && !claimed) {
                if (sumOld(O, O.pj[j], O.dp[j], O.tg[j]) === gU[g]) { nCand++; candP = O.tg[j]; }
              }
            }
          }
          if (nCand === 1) {
            okg[g] = true; effP[g] = candP;
            nCl++; clO[nCl] = gKO[g]; clD[nCl] = gDp[g]; clP[nCl] = candP;
            nAlt++;
          }
        }
      }
    }
    function copyRow(L) { var r = out[L].slice(); nFin++; fin[nFin] = r; return r; }
    for (g = 1; g <= gi; g++) {
      while (i2 < gStart[g]) { copyRow(i2); i2++; }
      var ok = (m > 0) ? !!okg[g] : false;
      var L;
      if (ok) {
        nVer++;
        for (L = gStart[g]; L <= gStart[g] + gCnt[g] - 1; L++) {
          var r = copyRow(L);
          var nmL = normTxt(S(out[L][18]));
          var sL = 0;
          for (j = 1; j <= m; j++) {
            if (!used[j] && O.nm[j] === nmL && O.pj[j] === gKO[g] && O.tg[j] === effP[g] && O.dp[j] === gDp[g]) { sL += O.ak[j]; used[j] = true; }
          }
          r[21] = sL; r[22] = r[20] - sL;
        }
        for (j = 1; j <= m; j++) {
          if (!used[j] && O.pj[j] === gKO[g] && O.tg[j] === effP[g] && O.dp[j] === gDp[g]) {
            if (addExtra) {
              nFin++;
              var rowE = new Array(OUT_COLS + 1);
              for (c = 1; c <= OUT_COLS; c++) rowE[c] = gCtx[g][c];
              var rr = ltrim(trim(O.raw[j]));
              if (rr.charAt(0) !== CIRCLE) rr = '  ' + rr;
              rowE[18] = rr;
              if (O.fx[j].length > 0) rowE[19] = O.fx[j];
              rowE[20] = 0; rowE[21] = O.ak[j]; rowE[22] = -O.ak[j];
              fin[nFin] = rowE; flg[nFin] = 2; nExtra++;
            } else nLeft++;
          }
        }
      } else {
        for (L = gStart[g]; L <= gStart[g] + gCnt[g] - 1; L++) {
          copyRow(L);
          if (gCnt[g] > 1 && gU[g] !== 0) { flg[nFin] = 1; nUnvLines++; }
        }
        if (gU[g] !== 0) nUnv++;
      }
      i2 = gStart[g] + gCnt[g];
    }
    while (i2 <= n) { copyRow(i2); i2++; }

    if (m > 0 && addExtra && O.hasCtx) {
      for (j = 1; j <= m; j++) {
        var claimed2 = false, jj2;
        for (jj2 = 1; jj2 <= nCl; jj2++) if (clO[jj2] === O.pj[j] && clD[jj2] === O.dp[j] && clP[jj2] === O.tg[j]) { claimed2 = true; break; }
        if (!claimed2) {
          var inPol = false;
          for (g = 1; g <= gi; g++) if (gDp[g] === O.dp[j] && S(gCtx[g][7]) === O.pol[j]) { inPol = true; break; }
          if (inPol) {
            var blk = false;
            for (g = 1; g <= gi; g++) if (!okg[g] && gU[g] !== 0 && gKO[g] === O.pj[j] && gDp[g] === O.dp[j]) { blk = true; break; }
            if (blk) nBlk++;
            else {
              var td = lookupDept(depts, O.dp[j]);
              nFin++;
              var rw = new Array(OUT_COLS + 1);
              if (cfmDate.length > 0) rw[1] = cfmDate;
              if (O.acc[j].length > 0) rw[2] = O.acc[j]; else if (td.acc.length > 0) rw[2] = td.acc;
              if (td.code.length > 0) rw[3] = td.code;
              if (O.fld[j].length > 0) rw[4] = O.fld[j];
              if (O.sec[j].length > 0) rw[5] = O.sec[j];
              if (td.com.length > 0) rw[6] = td.com;
              rw[7] = O.pol[j]; rw[8] = O.unit[j]; rw[10] = O.py[j]; rw[11] = O.pyN[j];
              if (gubun.length > 0) rw[12] = gubun;
              if (O.accN[j].length > 0) rw[13] = O.accN[j]; else if (td.accNm.length > 0) rw[13] = td.accNm;
              rw[14] = O.dp[j]; rw[15] = O.pj[j]; rw[16] = O.tg[j]; rw[17] = O.tgN[j];
              var rr2 = ltrim(trim(O.raw[j]));
              if (rr2.charAt(0) !== CIRCLE) rr2 = '  ' + rr2;
              rw[18] = rr2;
              if (O.fx[j].length > 0) rw[19] = O.fx[j];
              rw[20] = 0; rw[21] = O.ak[j]; rw[22] = -O.ak[j];
              fin[nFin] = rw; flg[nFin] = 2; nWhole++;
            }
          }
        }
      }
    }

    if (lineMode) {
      fin = []; flg = []; nFin = 0;
      for (var q2 = 1; q2 <= n; q2++) {
        var r4 = out[q2].slice();
        r4[21] = lineK[q2]; r4[22] = r4[20] - r4[21];
        nFin++; fin[nFin] = r4;
      }
    }
    var rows = [], flags = [], sumU = 0;
    for (i = 1; i <= nFin; i++) {
      var rr3 = [];
      for (c = 1; c <= OUT_COLS; c++) rr3.push(fin[i][c]);
      rows.push(rr3); flags.push(flg[i] || 0);
      sumU += Number(fin[i][21]) || 0;
    }
    var nBad = 0;
    for (g = 1; g <= gi; g++) if (gSum[g] !== gBud[g]) nBad++;

    var msg = [];
    msg.push(['ok', '변환 완료: ' + rows.length + '행 (' + titleOut + ')']);
    if (!lineMode) msg.push([sumLines === sumDept ? 'ok' : 'bad', '산출근거 금액 합계 ' + fmt(sumLines) + ' / 원본 부서 합계 ' + fmt(sumDept) + (sumLines === sumDept ? ' (일치)' : ' (불일치! 원본 확인 필요)')]);
    if (nBad > 0 && !lineMode) msg.push(['warn', '통계목 합계와 산출근거 합계가 다른 항목: ' + nBad + '건']);
    if (lineMode) {
      msg.push(['info', '[기정액] 요구서의 줄별 기정액을 그대로 입력 (2026 자료 불필요)']);
    } else if (m > 0) {
      msg.push(['info', '[기정액] 2026 본예산과 정확히 일치한 통계목 ' + nVer + '건은 줄별로 입력']);
      if (nAlt > 0) msg.push(['info', '그중 ' + nAlt + '건은 2026과 통계목 번호가 달라 같은 세부사업·편성목에서 금액이 같은 통계목으로 연결']);
      if (nExtra > 0) msg.push(['info', '같은 통계목에서 2027에 없는 2026 줄 ' + nExtra + '개를 예산액 0원 줄로 추가 (파란색)']);
      if (nWhole > 0) msg.push(['info', '2027 요구서에 나오지 않는 2026 사업의 줄 ' + nWhole + '개를 예산액 0원 줄로 추가 (파란색, 표 맨 아래)']);
      if (nBlk > 0) msg.push(['info', '확인 안 된 통계목이 있는 세부사업의 2026 줄 ' + nBlk + '개는 이중 계산을 피하려고 제외']);
      if (nLeft > 0) msg.push(['info', '추가하지 않은 2026 줄 ' + nLeft + '개 (삭제된 기정 줄 추가: 아니오)']);
      if (nUnv > 0) msg.push(['warn', '2026 자료로 확인 안 된 통계목 ' + nUnv + '건은 통계목 기정액을 첫 줄에 입력 (여러 줄 ' + nUnvLines + '줄은 주황색)']);
    } else {
      msg.push(['warn', O.why + ' 기정액은 통계목 단위로 첫 줄에 입력했습니다 (여러 줄은 주황색)']);
    }
    if (lineMode) msg.push(['info', '이 요구서에는 변경·신규 줄만 나옵니다. 예산액 합계 ' + fmt(sumLines) + ' / 기정액 합계 ' + fmt(sumU) + ' / 증감 ' + fmt(sumLines - sumU) + ' (부서 전체: 예산액 ' + fmt(sumDept) + ' / 기정액 ' + fmt(sumDeptJ) + ')']);
    else msg.push([sumU === sumDeptJ ? 'ok' : 'info', '기정액 합계 ' + fmt(sumU) + ' / 원본 부서 기정액 ' + fmt(sumDeptJ) + ' (차이 ' + fmt(sumU - sumDeptJ) + ')']);
    if (missing.length > 0 && depts.length > 0) msg.push(['warn', "부서 정보표에 없는 부서: " + missing.join(', ') + ' (부서코드·위원회 등 빈칸)']);
    return { rows: rows, flags: flags, title: titleOut, msg: msg, missing: missing };
  }

  // ---------- 시트3(전체 자료) -> 시트4 ----------
  var HN = ['예산확정일자', '회계코드', '부서코드', '분야명', '부문명', '위원회명', '정책사업명', '단위사업명',
    '사업형태구분명', '편성목', '편성목명', '예산구분', '회계명', '부서명', '세부사업명', '통계목', '통계목명',
    '산출근거명', '산출근거식', '예산액', '기정액', '비교증감', '국고보조금', '균특보조금', '기금보조금',
    '특별교부세', '도비보조금', '특별조정교부금', '자체재원', '지방채', '지방소멸대응기금'];
  function convertFullWide(grid3) {
    var lastRow = lastRowOf(grid3);
    if (lastRow === 0) return { error: '전체자료(시트3)가 비어 있습니다.' };
    var lastCol = lastColOf(grid3);
    if (lastRow < 3 || lastCol < 30) return { error: '전체자료의 형식이 맞지 않습니다. (제목 2줄 + 자료, 약 70개 열)' };
    var cm = [], miss = [], k, i;
    for (k = 1; k <= 31; k++) { cm[k] = hCol(grid3, HN[k - 1], lastCol); if (cm[k] === 0) miss.push(HN[k - 1]); }
    if (miss.length > 0) return { error: '첫 줄에서 다음 제목을 찾지 못했습니다: ' + miss.join(', ') + '. 제목 줄이 1~2행에 그대로 붙여넣어졌는지 확인하세요.' };
    var cEtc = hCol(grid3, '기타', lastCol), cYr = hCol(grid3, '회계연도', lastCol);
    var rows = [], sumT = 0, nFund = 0, nEtc = 0, yr = '', maxCh = 0;
    for (i = 3; i <= lastRow; i++) {
      if (S(txtVal(cell(grid3, i, cm[14]))).length > 0 && isAmt(cell(grid3, i, cm[20]))) {
        var r = new Array(33);
        for (k = 1; k <= 31; k++) r[k] = k >= 20 ? amt(cell(grid3, i, cm[k])) : txtVal(cell(grid3, i, cm[k]));
        if (r[23] > 0) r[32] = '국고';
        else if (r[24] > 0) r[32] = '균특';
        else if (r[25] > 0) r[32] = '기금';
        else if (r[27] > 0) r[32] = '도비';
        else r[32] = '자체';
        sumT += r[20];
        var fs = 0;
        for (k = 23; k <= 31; k++) fs += r[k];
        if (cEtc > 0 && amt(cell(grid3, i, cEtc)) !== 0) nEtc++;
        if (fs !== r[20]) nFund++;
        if (yr.length === 0 && cYr > 0) yr = S(txtVal(cell(grid3, i, cYr)));
        var s = S(r[12]);
        if (s.substring(0, 2) === '추경' && s.length >= 3 && /\d/.test(s.charAt(2))) {
          var ch = parseInt(s.charAt(2), 10);
          if (ch > maxCh) maxCh = ch;
        }
        rows.push(r.slice(1));
      }
    }
    if (rows.length === 0) return { error: '변환할 자료 행이 없습니다. (부서명과 예산액이 있는 줄이 필요합니다)' };
    var titleOut = maxCh > 0 ? '제' + maxCh + '회 추경 세출합본예산서(일반,특별)' : '본예산 세출합본예산서(일반,특별)';
    if (yr.length > 0) titleOut = yr + '년 ' + titleOut;
    var msg = [['ok', '변환 완료: ' + rows.length + '행 (' + titleOut + ')'], ['info', '예산액 합계 ' + fmt(sumT) + '천원']];
    if (nFund > 0) msg.push(['warn', '재원별 금액 합이 예산액과 다른 줄: ' + nFund + '건 (원본 확인)']);
    if (nEtc > 0) msg.push(['warn', "'기타' 재원 금액이 있는 줄 " + nEtc + '건: 출력에는 해당 열이 없어 제외됨']);
    return { rows: rows, flags: rows.map(function () { return 0; }), title: titleOut, msg: msg, missing: [] };
  }


  // ---------- 합본예산서(15100) 보고서형 파일 -> 합본예산서 ----------
  function numD(x) {
    var t = String(x == null ? '' : x).replace(/[\s,\r\n]/g, '');
    if (!t) return 0;
    var neg = false;
    if (/^\(.*\)$/.test(t)) t = t.slice(1, -1);
    if (t.charAt(0) === '△' || t.charAt(0) === '-') { neg = true; t = t.slice(1); }
    var v = parseFloat(t);
    return isNaN(v) ? 0 : (neg ? -v : v);
  }
  var FUND_COL = { '국': 23, '균': 24, '기': 25, '도': 27, '군': 29, '대': 31 };
  var FUND_RE = /^\s*\(?\s*([^\d\s(),△-])\s*([\d,]+)\s*\)?\s*$/;
  function isDetailGrid(grid) {
    for (var r = 1; r <= Math.min(grid.length, 15); r++) {
      var a = cell(grid, r, 1);
      if (!isEmpty(a) && /사업명세서|사업 명 세 서/.test(String(a).replace(/\s/g, '')) ) return true;
      if (!isEmpty(a) && trim(String(a)) === '부서:' ) return true;
    }
    return false;
  }
  function groupLines(lines) {
    var groups = {};
    lines.forEach(function (o, idx) { o.idx = idx; var key = [o.dept, o.pol, o.unit, o.proj, o.pcode, o.tcode].join('|'); (groups[key] = groups[key] || []).push(o); });
    Object.keys(groups).forEach(function (key) {
      var arr = groups[key];
      // 하위 줄이 딸린 윗줄(산출식 없음)은 금액을 0으로 볼 줄
      for (var j = 0; j < arr.length - 1; j++) {
        if (arr[j].formula.length === 0 && arr[j].block === arr[j + 1].block && arr[j + 1].lead > arr[j].lead) arr[j].parent = true;
      }
    });
    return groups;
  }
  function parseDetail(grid, depts) {
    var joiner = makeJoiner(grid, 2, 8);
    var lastRow = lastRowOf(grid), i, c;
    var dept = '', pol = '', unit = '', proj = '', pcode = '', pname = '', tcode = '', tname = '', block = '', last = null;
    var lines = [], missing = [], unknownLetters = 0;
    var yr = '';
    for (i = 1; i <= 6; i++) { var tt = cell(grid, i, 1); if (!isEmpty(tt)) { var ym = String(tt).match(/(\d{4})\s*년/); if (ym) { yr = ym[1]; break; } } }
    var dctx = { found: false, code: '', com: '', acc: '', accNm: '' };
    function rs() { unit = proj = pcode = pname = tcode = tname = block = ''; last = null; }
    for (i = 1; i <= lastRow; i++) {
      var c1 = trim(joinWrap(cell(grid, i, 1), ''));
      if (c1.length) {
        if (/^- \d+ -$/.test(c1)) continue;
        if (c1 === '부서:' || c1 === '정책:' || c1 === '단위:' || c1.indexOf('부서ㆍ') === 0) continue;
        if (trim(S(cell(grid, i, 9))).length && isAmt(cell(grid, i, 9))) {
          dept = c1; pol = ''; rs();
          dctx = lookupDept(depts, dept);
          if (!dctx.found && missing.indexOf(dept) < 0) missing.push(dept);
        }
        continue;
      }
      var t2 = trim(joinWrap(cell(grid, i, 2), '')), t3 = trim(joinWrap(cell(grid, i, 3), '')), t4 = trim(joinWrap(cell(grid, i, 4), ''));
      if (t2.length) { pol = t2; rs(); continue; }
      if (t3.length) { unit = t3; proj = pcode = pname = tcode = tname = block = ''; last = null; continue; }
      if (t4.length) { proj = t4; pcode = pname = tcode = tname = block = ''; last = null; continue; }
      var t6 = trim(joinWrap(cell(grid, i, 6), ''));
      if (t6.length) {
        var m6 = t6.match(/^(\d{3})\s*(.*)$/);
        if (m6) { pcode = m6[1]; pname = m6[2]; }
        tcode = tname = block = ''; last = null; continue;
      }
      var raw7 = cell(grid, i, 7);
      var t7 = trim(joinWrap(raw7, ''));
      if (t7.length) {
        if (/^(본예산|추경\s*\d+\s*회)$/.test(t7)) { block = t7.replace(/\s/g, ''); last = null; continue; }
        var m7 = t7.match(/^(\d\d)\s+(.+)$/);
        if (m7 && !/^\s/.test(String(raw7))) { tcode = m7[1]; tname = m7[2]; block = ''; last = null; continue; }
        var txt = joiner(raw7);
        var lead = (txt.match(/^ */) || [''])[0].length;
        var o = { dept: dept, pol: pol, unit: unit, proj: proj, pcode: pcode, pname: pname, tcode: tcode, tname: tname, block: block, ctx: dctx,
          text: txt.replace(/^ {2}/, ''), lead: lead, formula: trim(joinWrap(cell(grid, i, 8), '')),
          b: numD(cell(grid, i, 9)), k: numD(cell(grid, i, 10)), dRaw: trim(S(cell(grid, i, 11)).replace(/[\r\n]/g, '')), f: {}, hasF: false, dead: false, parent: false };
        o.d = o.dRaw.length ? numD(cell(grid, i, 11)) : o.b - o.k;
        lines.push(o); last = o; continue;
      }
      var fm = S(cell(grid, i, 9)).match(FUND_RE);
      if (fm) {
        if (last) {
          var fc = FUND_COL[fm[1]];
          if (fc) { last.f[fc] = (last.f[fc] || 0) + numD(fm[2]); last.hasF = true; } else unknownLetters++;
        }
      }
    }
    var groups0 = groupLines(lines);
    return { lines: lines, missing: missing, yr: yr, unknownLetters: unknownLetters };
  }
  function convertDetail(grid, set) {
    set = set || {};
    var depts = set.depts || [], dates = set.dates || {}, cfmDate = trim(S(set.cfmDate)), accts = set.accts || DEFAULT_ACCTS;
    var P = parseDetail(grid, depts), lines = P.lines, missing = P.missing, yr = P.yr, unknownLetters = P.unknownLetters, c;
    if (lines.length === 0) return { error: '변환할 줄을 찾지 못했습니다. (합본예산서(15100)에서 저장한 파일인지 확인하세요)' };

    function rank(b) { return b === '본예산' ? 0 : parseInt(b.replace(/\D/g, ''), 10) || 0; }
    var groups = groupLines(lines);
    var nSuper = 0, nSuperAmt = 0;
    Object.keys(groups).forEach(function (key) {
      var arr = groups[key], j;
      for (j = 0; j < arr.length; j++) {
        var X = arr[j], rx = rank(X.block);
        if (rx === 0) continue;
        var nm = X.text.replace(/\s/g, ''), best = null, q, bestAmt = null;
        for (q = 0; q < j; q++) {
          var Y = arr[q];
          if (!Y.dead && rank(Y.block) < rx && Y.text.replace(/\s/g, '') === nm && Y.lead === X.lead) {
            if (!best || rank(Y.block) > rank(best.block)) best = Y;
            if (X.k !== 0 && Y.b === X.k && (!bestAmt || rank(Y.block) > rank(bestAmt.block))) bestAmt = Y;
          }
        }
        if (bestAmt) best = bestAmt;
        if (best) { best.dead = true; nSuper++; continue; }
        if (X.k !== 0) {
          for (q = 0; q < j; q++) {
            var Z = arr[q];
            if (!Z.dead && rank(Z.block) < rx && Z.lead === X.lead && Z.b === X.k) { Z.dead = true; nSuperAmt++; break; }
          }
        }
      }
    });

    var rows = [], maxCh = 0, sumT = 0, nDate = {};
    lines.forEach(function (o) {
      if (o.dead) return;
      var sp = splitPolicy(o.pol);
      var r = new Array(OUT_COLS);
      var dt = dates[o.block] !== undefined && S(dates[o.block]).length ? S(dates[o.block]) : cfmDate;
      var ua = splitUnitAcct(o.unit, accts);
      r[0] = opt(dt); r[1] = opt(ua ? ua.acc : o.ctx.acc); r[2] = opt(o.ctx.code); r[3] = opt(sp.fld); r[4] = opt(sp.sec); r[5] = opt(o.ctx.com);
      r[6] = sp.nm; r[7] = ua ? ua.unit : o.unit;
      r[9] = o.pcode; r[10] = o.pname; r[11] = o.block; r[12] = opt(ua ? ua.accNm : o.ctx.accNm); r[13] = o.dept; r[14] = o.proj;
      r[15] = o.pcode + '-' + o.tcode; r[16] = o.tname; r[17] = o.text;
      if (o.formula.length) r[18] = o.formula;
      var b = o.parent ? 0 : o.b, d = o.parent ? 0 : o.d;
      r[19] = b; r[20] = o.k; r[21] = d;
      for (c = 23; c <= 31; c++) r[c - 1] = 0;
      if (!o.parent) {
        if (o.hasF) for (c = 23; c <= 31; c++) r[c - 1] = o.f[c] || 0;
        else r[28] = b;
      }
      if (r[22] > 0) r[31] = '국고'; else if (r[23] > 0) r[31] = '균특'; else if (r[24] > 0) r[31] = '기금'; else if (r[26] > 0) r[31] = '도비'; else r[31] = '자체';
      var rk = rank(o.block); if (rk > maxCh) maxCh = rk;
      if (!dates[o.block] && !nDate[o.block]) nDate[o.block] = true;
      sumT += b;
      rows.push(r);
    });
    var titleOut = maxCh > 0 ? '제' + maxCh + '회 추경 세출합본예산서(일반,특별)' : '본예산 세출합본예산서(일반,특별)';
    if (yr) titleOut = yr + '년 ' + titleOut;
    var msg = [['ok', '변환 완료: ' + rows.length + '행 (' + titleOut + ')'], ['info', '예산액 합계 ' + fmt(sumT) + '천원']];
    if (nSuper + nSuperAmt > 0) msg.push(['info', '앞 회차의 같은 줄을 대체한 ' + (nSuper + nSuperAmt) + '줄은 최신 내용만 남기고 이전 줄은 제외했습니다 (이름이 같은 줄 ' + nSuper + '건, 기정액과 금액이 같은 줄 ' + nSuperAmt + '건)']);
    msg.push(['info', '재원별 금액은 각 줄 아래의 국·균·기·도·군·대 금액을 사용하고, 없으면 자체재원(군)으로 처리했습니다.']);
    if (unknownLetters > 0) msg.push(['warn', '알 수 없는 재원 글자가 있는 줄 ' + unknownLetters + '개는 재원 금액에서 제외됨']);
    var noDate = Object.keys(nDate);
    if (noDate.length && Object.keys(dates).length) msg.push(['info', '예산확정일자를 입력하지 않은 회차: ' + noDate.join(', ') + ' (해당 열은 빈칸)']);
    if (missing.length && depts.length) msg.push(['warn', '부서 정보표에 없는 부서: ' + missing.join(', ') + ' (부서코드·위원회·회계 빈칸)']);
    msg.push(['info', '산출근거명이 줄바꿈으로 잘려 있던 곳의 띄어쓰기는 추정한 것이라 일부 다를 수 있습니다.']);
    return { rows: rows, flags: rows.map(function () { return 0; }), title: titleOut, msg: msg, missing: missing };
  }

  function convertFull(grid3, set) {
    if (!grid3 || lastRowOf(grid3) === 0) return { error: '전체자료(시트3)가 비어 있습니다.' };
    if (isDetailGrid(grid3) && lastColOf(grid3) < 30) return convertDetail(grid3, set);
    return convertFullWide(grid3);
  }

  // ---------- 붙여넣은 글자(탭 구분) -> 격자 ----------
  function parseTSV(text) {
    var rows = [], row = [], f = '', i = 0, n = text.length, inQ = false, fresh = true;
    while (i < n) {
      var ch = text.charAt(i);
      if (inQ) {
        if (ch === '"') {
          if (text.charAt(i + 1) === '"') { f += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        f += ch; i++; continue;
      }
      if (ch === '"' && fresh) { inQ = true; fresh = false; i++; continue; }
      if (ch === '\t') { row.push(f); f = ''; fresh = true; i++; continue; }
      if (ch === '\r' && text.charAt(i + 1) === '\n') { i++; continue; }
      if (ch === '\n' || ch === '\r') { row.push(f); rows.push(row); row = []; f = ''; fresh = true; i++; continue; }
      f += ch; fresh = false; i++;
    }
    if (f.length > 0 || row.length > 0) { row.push(f); rows.push(row); }
    return rows;
  }

  // ---------- 엑셀 파일 만들기 (ExcelJS) ----------
  var CENTER = { 2: 1, 3: 1, 10: 1, 12: 1, 16: 1, 32: 1 };
  var TEXTFMT = { 1: 1, 2: 1, 3: 1, 10: 1, 16: 1, 19: 1 };
  function colName(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  async function buildWorkbook(ExcelJS, templateBuf, result) {
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuf);
    var ws = wb.worksheets[0];
    var rows = result.rows, n = rows.length, last = 4 + n;
    var thin = { style: 'thin' };
    var border = { top: thin, left: thin, bottom: thin, right: thin };
    var sums = [];
    for (var ci = 20; ci <= 31; ci++) sums[ci] = 0;
    for (var i = 0; i < n; i++) {
      var row = ws.getRow(5 + i), r = rows[i];
      for (var c = 1; c <= OUT_COLS; c++) {
        var cl = row.getCell(c), v = r[c - 1];
        if (v !== undefined && v !== null && v !== '') cl.value = v;
        cl.font = { name: '맑은 고딕', size: 11 };
        cl.border = border;
        if (c >= 20 && c <= 31) {
          cl.numFmt = '#,##0'; cl.alignment = { horizontal: 'right' };
          sums[c] += Number(v) || 0;
        } else {
          cl.alignment = { horizontal: CENTER[c] ? 'center' : 'left' };
          if (TEXTFMT[c]) cl.numFmt = '@';
        }
      }
      var fl = result.flags[i];
      if (fl === 1) for (c = 21; c <= 22; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
      if (fl === 2) for (c = 18; c <= 22; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
      row.commit && row.commit();
    }
    for (c = 20; c <= 31; c++) {
      var cn = colName(c);
      ws.getCell(cn + '2').value = { formula: 'SUBTOTAL(9,' + cn + '5:' + cn + last + ')', result: sums[c] };
    }
    ws.getCell('L1').value = result.title;
    ws.autoFilter = 'A4:AF' + last;
    return wb.xlsx.writeBuffer();
  }

  var api = { convertDetail: convertDetail, isDetailGrid: isDetailGrid, convertReport: convertReport, convertFull: convertFull, parseTSV: parseTSV, buildWorkbook: buildWorkbook, lastRowOf: lastRowOf, lastColOf: lastColOf, hCol: hCol };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.BudgetConv = api;
})(typeof window !== 'undefined' ? window : globalThis);
