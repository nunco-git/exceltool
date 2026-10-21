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
    var s = trim(String(x)).replace(/,/g, '');
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

  // ---------- 2026 자료(시트3) 읽기 ----------
  function loadOld(grid3) {
    var o = { m: 0, why: '', hasCtx: false };
    if (!grid3) { o.why = '2026 전체자료가 입력되지 않아 줄별 기정액을 확인하지 못했습니다.'; return o; }
    var lastRow = lastRowOf(grid3), lastCol = lastColOf(grid3);
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
  function sumOld(O, o, d, p) {
    var t = 0;
    for (var j = 1; j <= O.m; j++) if (O.pj[j] === o && O.dp[j] === d && O.tg[j] === p) t += O.ak[j];
    return t;
  }

  // ---------- 시트1(예산요구서) -> 시트2 ----------
  function convertReport(grid1, grid3, set) {
    set = set || {};
    var depts = set.depts || [];
    var cfmDate = trim(S(set.cfmDate));
    var gubunSet = trim(S(set.gubun));
    var addExtra = set.addExtra !== false;
    var lastRow = lastRowOf(grid1);
    if (lastRow === 0) return { error: '예산요구서(시트1)가 비어 있습니다.' };
    var v = function (i, c) { return cell(grid1, i, c); };
    var i, k, c, g, j, a, b, cc, d;

    var titleSrc = '';
    for (i = 1; i <= 8; i++) {
      if (i > lastRow) break;
      a = joinWrap(v(i, 1), ' ');
      if (a.indexOf('년') >= 0 && a.indexOf('예산') >= 0) { titleSrc = trim(a); break; }
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
          dept = a; pol = ''; fld = ''; sec = ''; unit = ''; proj = '';
          pyeong = ''; pname = ''; tong = ''; tname = '';
          sumDept += amt(v(i, 9));
          sumDeptJ += amt(v(i, 10));
          var dd = lookupDept(depts, dept);
          dCode = dd.code; dCom = dd.com; dAcc = dd.acc; dAccNm = dd.accNm;
          if (!dd.found && missing.indexOf(dept) < 0) missing.push(dept);
        }
      } else if (b.length > 0) {
        var sp = splitPolicy(b);
        pol = sp.nm; fld = sp.fld; sec = sp.sec;
      } else if (cc.length > 0) {
        unit = cc;
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
            gBud[gi] = amt(v(i, 9)); gU[gi] = amt(v(i, 10));
            gKO[gi] = proj; gKP[gi] = tong; gDp[gi] = dept;
            var cx = [];
            cx[1] = opt(cfmDate); cx[2] = opt(dAcc); cx[3] = opt(dCode); cx[4] = opt(fld); cx[5] = opt(sec); cx[6] = opt(dCom);
            cx[7] = pol; cx[8] = unit; cx[10] = pyeong; cx[11] = pname; cx[12] = opt(gubun); cx[13] = opt(dAccNm);
            cx[14] = dept; cx[15] = proj; cx[16] = tong; cx[17] = tname;
            gCtx[gi] = cx;
          }
        } else {
          var txt = rtrim(joinWrap(v(i, 6), ''));
          var st = ltrim(txt);
          txt = st.charAt(0) === CIRCLE ? st : '  ' + st;
          var t = amt(v(i, 9));
          var u = 0;
          if (gi > 0 && gCnt[gi] === 0) u = gU[gi];
          n++;
          var o = new Array(OUT_COLS + 1);
          o[1] = opt(cfmDate); o[2] = opt(dAcc); o[3] = opt(dCode); o[4] = opt(fld); o[5] = opt(sec); o[6] = opt(dCom);
          o[7] = pol; o[8] = unit; o[10] = pyeong; o[11] = pname; o[12] = opt(gubun); o[13] = opt(dAccNm);
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
    var O = loadOld(grid3);
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
    msg.push([sumLines === sumDept ? 'ok' : 'bad', '산출근거 금액 합계 ' + fmt(sumLines) + ' / 원본 부서 합계 ' + fmt(sumDept) + (sumLines === sumDept ? ' (일치)' : ' (불일치! 원본 확인 필요)')]);
    if (nBad > 0) msg.push(['warn', '통계목 합계와 산출근거 합계가 다른 항목: ' + nBad + '건']);
    if (m > 0) {
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
    msg.push([sumU === sumDeptJ ? 'ok' : 'info', '기정액 합계 ' + fmt(sumU) + ' / 원본 부서 기정액 ' + fmt(sumDeptJ) + ' (차이 ' + fmt(sumU - sumDeptJ) + ')']);
    if (missing.length > 0) msg.push(['warn', "부서 정보표에 없는 부서: " + missing.join(', ') + ' (부서코드·위원회 등 빈칸)']);
    if (cfmDate.length === 0) msg.push(['info', '예산확정일자를 입력하면 해당 열이 채워집니다.']);
    return { rows: rows, flags: flags, title: titleOut, msg: msg, missing: missing };
  }

  // ---------- 시트3(전체 자료) -> 시트4 ----------
  var HN = ['예산확정일자', '회계코드', '부서코드', '분야명', '부문명', '위원회명', '정책사업명', '단위사업명',
    '사업형태구분명', '편성목', '편성목명', '예산구분', '회계명', '부서명', '세부사업명', '통계목', '통계목명',
    '산출근거명', '산출근거식', '예산액', '기정액', '비교증감', '국고보조금', '균특보조금', '기금보조금',
    '특별교부세', '도비보조금', '특별조정교부금', '자체재원', '지방채', '지방소멸대응기금'];
  function convertFull(grid3) {
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

  var api = { convertReport: convertReport, convertFull: convertFull, parseTSV: parseTSV, buildWorkbook: buildWorkbook, lastRowOf: lastRowOf, lastColOf: lastColOf, hCol: hCol };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.BudgetConv = api;
})(typeof window !== 'undefined' ? window : globalThis);
