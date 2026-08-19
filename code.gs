// ════════════════════════════════════════════════════════════════════
//  WABI NAVI — Google Apps Script v8.1
//  注文記録 + 会員ポイント管理 + 予約管理 + ホームページ連動
//  + 🍳 キッチン表示システム（KDS）／注文シートを新着順に
// ════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════
//  🌐 サイトURL設定（ドメイン移行はここだけ変更すればOK）
// ════════════════════════════════════════════════════════════════════
//  SITE_URL … ホームページ。WordPress解約→Cloudflareへ切替後に
//              "https://wabinavi.com" へ変更してください。
//  APP_BASE … 会員・注文アプリ。★印刷済みQRコードが指しているため変更禁止★
const SITE_URL = "https://riichi-ramen.github.io/";           // ← 移行後に書き換える
const APP_BASE = "https://riichi-ramen.github.io/wabinavi-order/"; // 変更禁止

const SHEET_ORDERS  = "Zamówienia / Orders";
const SHEET_MEMBERS = "Członkowie / Members";
// ★★★ 係員パスワード ★★★
// ⚠️ 2026-08-09：この値は開発チャット上に露出したため、変更を強く推奨します。
//    変更手順：下の1行を書き換え → デプロイを管理 → 新しいバージョン（全ページに即反映）
const STAFF_PASSWORD = "TanTanMen"; // ← 係員パスワード（設定必須）

// ── ランク定義 ──────────────────────────────────────────────────────
const RANKS = [
  { name:"Sakura", icon:"🌸", minPt:0,   maxPt:49  },
  { name:"Ume",    icon:"🏅", minPt:50,  maxPt:99  },
  { name:"Take",   icon:"🥉", minPt:100, maxPt:199 },
  { name:"Matsu",  icon:"🥈", minPt:200, maxPt:499 },
  { name:"Fuji",   icon:"👑", minPt:500, maxPt:99999 },
];

function getRank(pt) {
  return RANKS.slice().reverse().find(r => pt >= r.minPt) || RANKS[0];
}

// ════════════════════════════════════════════════════════════════════
//  🎫 特典クーポン制度（確定仕様 2026-07-11）
//  1杯=10pt。累計が50ptの倍数に到達するたびにクーポン獲得（交互）：
//    奇数×50pt（50, 150, 250…）→ ラーメン半額クーポン
//    100の倍数（100, 200, 300…）→ ラーメン1杯無料クーポン
//  ※半額・無料で提供した丼にはポイントを付与しない
//  ※クーポンは会員シートR列に記録。注文で使用すると使用済みになる
//  ※獲得時に自動メール通知。次回注文の確認画面に自動表示される
// ════════════════════════════════════════════════════════════════════
const MEMBER_COUPON_COL = 18; // R列

// ── 係員パスワード照合＋総当たり対策（v5.7） ─────────────────────
//  5回連続で間違えると15分間ロック。正しい入力で失敗カウントはリセット。
//  これにより短めのパスワードでも総当たりが実質不可能になる
//  （6桁でも 5回/15分 では全探索に数百年かかる）。
//  注：ロックは全端末共通。第三者がわざと連続失敗させると正規スタッフも
//  15分待ちになるが、その間もお客様のスタンプ等（パスワード不要機能）は動く。
function verifyStaffPassword(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("staff_pw_fails") || 0);
  if (fails >= 5) return "locked";
  if ((pw || "") === STAFF_PASSWORD) {
    if (fails) { try { cache.remove("staff_pw_fails"); } catch (e) {} }
    return "ok";
  }
  try { cache.put("staff_pw_fails", String(fails + 1), 900); } catch (e) {}
  return "wrong";
}
const STAFF_AUTH_MSG = {
  locked: "パスワード試行回数の上限に達しました。15分後に再試行してください。/ Zbyt wiele prób — spróbuj za 15 minut.",
  wrong:  "パスワードが違います / Nieprawidłowe hasło"
};

// ── 会員行の照合ヘルパー（v5.4） ─────────────────────────────────
//  会員シートは A列=会員ID / B列=連絡先（登録時のメール or 電話）。
//  queryが会員ID・メール・電話のどれでも一致を判定できる。
//  電話は +48 やスペースの表記ゆれを吸収するため数字末尾9桁でも照合。
function memberRowMatches(rowId, rowContact, query) {
  const q = (query || "").toString().trim().toLowerCase();
  if (!q) return false;
  const a = (rowId      || "").toString().trim().toLowerCase();
  const b = (rowContact || "").toString().trim().toLowerCase();
  if (q === a || q === b) return true;
  const qd = q.replace(/\D/g, "").slice(-9);
  const bd = b.replace(/\D/g, "").slice(-9);
  return qd.length === 9 && bd.length === 9 && qd === bd;
}

function awardMilestoneCoupons(email, name, oldPt, newPt) {
  const earned = [];
  for (let m = Math.floor(oldPt / 50) * 50 + 50; m <= newPt; m += 50) {
    earned.push({
      t: (m % 100 === 0) ? "free" : "half",
      m: m,
      e: Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd"),
      u: null
    });
  }
  if (!earned.length) return null;
  try {
    const coupons = getMemberCoupons(email);
    earned.forEach(c => coupons.push(c));
    saveMemberCoupons(email, coupons);
    earned.forEach(c => sendCouponMail(email, name, c));
  } catch (err) { Logger.log("awardMilestoneCoupons: " + err); }
  return earned;
}

function getMemberCoupons(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getMemberSheet(ss);
  const data = sh.getDataRange().getValues();
  const key = (email || "").toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (memberRowMatches(data[i][0], data[i][1], key)) {
      const raw = (data[i][MEMBER_COUPON_COL - 1] || "").toString().trim();
      if (!raw) return [];
      try { return JSON.parse(raw); } catch (e) { return []; }
    }
  }
  return [];
}

function saveMemberCoupons(email, coupons) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getMemberSheet(ss);
  const data = sh.getDataRange().getValues();
  const key = (email || "").toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (memberRowMatches(data[i][0], data[i][1], key)) {
      sh.getRange(i + 1, MEMBER_COUPON_COL).setValue(JSON.stringify(coupons));
      return;
    }
  }
}

function getUnusedCoupons(email) {
  return getMemberCoupons(email).filter(c => !c.u);
}

// 会員シートにクーポン列（R列）のヘッダーを追加：1回実行
function setupCouponsColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getMemberSheet(ss);
  sh.getRange(1, MEMBER_COUPON_COL).setValue("クーポン / Kupony")
    .setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e").setFontSize(9);
  sh.setColumnWidth(MEMBER_COUPON_COL, 220);
  Logger.log("✅ クーポン列（R列）を設定しました");
}

function sendCouponMail(email, name, coupon) {
  if (!email) return;
  const isFree = coupon.t === "free";
  const couponName = isFree
    ? "🍜 DARMOWY Ramen / ラーメン1杯 無料クーポン"
    : "🎫 50% ZNIŻKI na Ramen / ラーメン半額クーポン";
  const subject = "🎉 Wabi Navi — Zdobyłeś kupon! / クーポン獲得！（" + coupon.m + "pt達成）";
  const body = `
${name ? name + " 様 / Drogi/a " + name + "," : "Drogi Członku,"}

━━━━━━━━━━━━━━━━━━━━━━━━━━
🎉 Gratulacje! / おめでとうございます！
Twoje punkty osiągnęły ${coupon.m} pt!
累計ポイントが ${coupon.m}pt に到達しました！
━━━━━━━━━━━━━━━━━━━━━━━━━━

Otrzymujesz / 獲得クーポン:
${couponName}

📱 Jak użyć / 使い方:
Kupon pojawi się automatycznie na ekranie potwierdzenia zamówienia
podczas następnej wizyty — wystarczy go zaznaczyć.
次回ご来店時、注文確認画面に自動表示されます。選択するだけでOKです。
Zniżka zostanie naliczona przy kasie.
割引はレジでのお会計時に適用されます。

[Zasady / ルール]
・Przy wizycie z użyciem kuponu pieczątka (punkty) nie jest przyznawana.
　クーポン（半額・無料）をご使用の来店は、スタンプ（ポイント）対象外です。
・Kupony: co 50 pkt na przemian — 50pt=−50%, 100pt=gratis, 150pt=−50%, 200pt=gratis…
　50ptごとに半額と無料が交互に貯まります。
・Nie łączy się ze zniżką studencką/seniorską.
　学割・シニア割との併用はできません。

Do zobaczenia! / またのご来店をお待ちしております！
🍜 Wabi Navi — Toruń

━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI Members Club
━━━━━━━━━━━━━━━━━━━━━━━━━━`;
  try { GmailApp.sendEmail(email, subject, body.trim()); }
  catch (e) { Logger.log("sendCouponMail: " + e); }
}

// ── doPost：注文記録 + スタンプ処理 ──────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || "order";

    if (action === "order")           return handleOrder(data);
    if (action === "stamp")           return handleStamp(data);
    if (action === "register")        return handleRegister(data);
    if (action === "getCard")         return handleGetCard(data);
    if (action === "staffAddPoint")   return handleStaffAddPoint(data);
    if (action === "staffRemovePoint") return handleStaffRemovePoint(data);
    if (action === "staffQrStamp")    return handleStaffQrStamp(data);
    if (action === "submitQuiz")      return handleQuizSubmit(data);
    if (action === "submitTrivia")    return handleTriviaAnswer(data);
    if (action === "withdraw")        return handleWithdraw(data);
    if (action === "setHiddenItems")  return handleSetHiddenItems(data);
    if (action === "borrowBook")      return handleBorrowBook(data);
    if (action === "cancelReservation")  return handleCancelReservation(data);
    if (action === "updateReservation")  return handleUpdateReservation(data);
    if (action === "staffReservation")   return handleStaffReservation(data);
    if (action === "checkout")           return handleCheckout(data);
    if (action === "webReservation")     return handleWebReservation(data);
    if (action === "kitchenUpdate")      return handleKitchenUpdate(data);

    return jsonResponse({ status:"error", message:"Unknown action" });
  } catch(err) {
    return jsonResponse({ status:"error", message: err.toString() });
  }
}

// ── doGet：会員カード取得（GETリクエスト対応）────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || "";
    const email  = e.parameter.email  || "";

    if (action === "getBooks") {
      return jsonResponse(getBooksList());
    }
    if (action === "getMenu") {
      const m = getMenuFromSheet();
      return jsonResponse(m || { status:"noSheet",
        message:"メニューシート未作成。setupMenuSheet()を実行してください。" });
    }
    if (action === "getTrivia") {
      return jsonResponse(getTriviaForMember(e.parameter.memberId));
    }
    if (action === "getLessons") {
      return jsonResponse(getLessonsPublic());
    }
    if (action === "getHiddenItems") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      return jsonResponse({ status:"ok", hidden: getHiddenItemIds(ss) });
    }
    if (action === "getCard" && email) {
      const member = findMember(email);
      if (!member) return jsonResponse({ status:"notFound" });
      member.coupons = getUnusedCoupons(email);
      return jsonResponse({ status:"ok", member });
    }
    if (action === "getReservation") {
      return jsonResponse(getReservationPublic(e.parameter.id, e.parameter.t, e.parameter.email));
    }
    if (action === "getReservations") {
      return jsonResponse(getReservationsForStaff(e.parameter.password, e.parameter.range));
    }
    if (action === "lookupNip") return jsonResponse(lookupNip(e.parameter.nip));
    if (action === "getKitchen") {
      return jsonResponse(getKitchenData(e.parameter.password,
                                         e.parameter.mode,
                                         e.parameter.date));
    }
    if (action === "checkMember") {
      return jsonResponse(checkMemberPublic(e.parameter.email));
    }
    // ── ヘルスチェック：ブラウザでURLを直接開くと状態を表示 ──
    return jsonResponse({
      status: "ok",
      app: "WABI NAVI GAS",
      version: "v8.1-kds",
      time: Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd HH:mm:ss"),
      message: "✅ デプロイ正常 / Deployment OK — このJSONが見えていれば接続設定は正しいです"
    });
  } catch(err) {
    return jsonResponse({ status:"error", message: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
//  注文記録処理
// ════════════════════════════════════════════════════════════════════
function handleOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) sh = createOrderSheet(ss);

  const orderId = data.orderId || "";
  const time    = new Date(data.timestamp);
  const stolik  = "Stolik / Table " + data.table;
  const goscie  = data.guests + " os. / guests";
  let uwagi     = data.uwagi || "";
  const total   = data.total || 0;
  const items   = data.items || [];

  // 🎫 クーポン使用処理（会員のみ・該当種の未使用クーポンを1枚消費）
  let couponUsed = null;
  if (data.useCoupon && data.isMember && data.memberContact) {
    const cEmail = data.memberContact.trim().toLowerCase();
    const coupons = getMemberCoupons(cEmail);
    const ci = coupons.findIndex(c => !c.u && c.t === data.useCoupon);
    if (ci >= 0) {
      coupons[ci].u = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
      saveMemberCoupons(cEmail, coupons);
      couponUsed = data.useCoupon;
      const dAmt = Number(data.couponDiscount) || 0;
      const dTgt = data.couponTarget || "";
      uwagi = "⚠️KONTROLA KASY/照合注意 "
        + (couponUsed === "free"
        ? "【🎫 KUPON: DARMOWY RAMEN" : "【🎫 KUPON: 50% ZNIŻKI")
        + (dTgt ? " — " + dTgt : "") + (dAmt ? " −" + dAmt + " zł" : "")
        + "（zniżka już w sumie / 割引は合計に反映済み）】"
        + (uwagi ? " " + uwagi : "");
    } else {
      uwagi = "【⚠️ KUPON NIEZWERYFIKOWANY — sprawdź przy kasie / クーポン確認不可・レジで確認】"
        + (uwagi ? " " + uwagi : "");
    }
  }
  const czlonek = data.isMember ? "✅ Członek / Member" : "Gość / Guest";

  // v8.0：末尾に足すのではなく2行目に差し込む（最新の注文が常に一番上）
  const rowsOut = items.map((item, idx) => {
    const memoLines = [];
    const customs   = item.customizations || "";
    if (customs) {
      customs.split(" | ").forEach(part => {
        part = part.trim();
        if (!part) return;
        if (part.startsWith("Bez:")) {
          part.replace("Bez:", "").trim().split(",").forEach(ing => {
            ing = ing.trim();
            if (ing) memoLines.push("🚫 POMIŃ / REMOVE: " + ing);
          });
        } else {
          memoLines.push("🔧 ZMIANA / CHANGE: " + part);
        }
      });
    }
    const memo = memoLines.length > 0 ? memoLines.join("\n") : "—";

    return [
      orderId, time, stolik, goscie, idx+1,
      item.namePL, item.nameEN, item.qty, item.unitPrice, item.lineTotal,
      memo,
      idx===0 ? uwagi        : "",
      idx===0 ? czlonek      : "",
      idx===0 ? (data.memberName    ||"") : "",
      idx===0 ? (data.memberContact ||"") : "",
      idx===0 ? (data.memberBirthMonth||"") : "",
      idx===0 ? (data.memberLanguage||"") : "",
      "NOWE / NEW",
      idx===0 ? total : "",
      item.id || ""                      // T列：商品ID（キッチン集計用・v8.0）
    ];
  });

  const startRow = prependOrderRows_(sh, rowsOut);

  formatOrderRows(sh, startRow, items.length);

  // 🍚 ライスセット割の表記
  const riceAmt = Number(data.riceSetDiscount) || 0;
  if (riceAmt > 0) {
    const riceNote = "【🍚 ZESTAW: ryż z ramenem ×" + (data.riceSetCount || 1) +
      " −" + riceAmt + " zł（合計反映済み）】";
    sh.getRange(startRow, 12).setValue(riceNote + " " + (sh.getRange(startRow, 12).getValue() || ""));
  }

  // 🎁 ランク特典（トッピング無料）の表記
  const perkAmt = Number(data.rankPerkDiscount) || 0;
  if (perkAmt > 0) {
    const perkNote = "【🎁 BONUS RANGI " + (data.rankPerkRank || "") + ": " +
      (data.rankPerkDetail || "dodatki") + " gratis −" + perkAmt + " zł（合計反映済み）】";
    sh.getRange(startRow, 12).setValue(perkNote + " " + (sh.getRange(startRow, 12).getValue() || ""));
  }

  // 🎫 クーポン・🎁特典の注文 → シート上で「照合注意」を強調表示
  if (couponUsed || perkAmt > 0) {
    const rowCount = Math.max(items.length, 1);
    // 注文ブロック全体を黄色でハイライト
    sh.getRange(startRow, 1, rowCount, 19).setBackground("#fff3cd");
    // UWAGI欄（クーポン表記入り）を赤太字に
    sh.getRange(startRow, 12).setFontColor("#a81f10").setFontWeight("bold").setFontSize(10);
    // 合計セルを赤背景＋照合メモ付きに
    sh.getRange(startRow, 19)
      .setBackground("#f8d7da").setFontColor("#a81f10").setFontWeight("bold")
      .setNote("⚠️ KONTROLA KASY / 照合注意\n" +
        "Kupon użyty — suma w arkuszu jest już PO zniżce.\n" +
        "Sprawdź zgodność z POS przed wydaniem paragonu.\n" +
        "クーポン使用のため、この合計は割引後の金額です。\n" +
        "レシート発行前にPOSの合計と一致するか確認してください。");
    // 1列目に警告マーカー
    sh.getRange(startRow, 1)
      .setValue("⚠️" + (sh.getRange(startRow, 1).getValue() || ""))
      .setFontColor("#a81f10").setFontWeight("bold");
  }

  // 会員登録が含まれる場合はシートに登録
  // v7.0修正：旧版は email/name/birthMonth/language の4項目しか渡しておらず、
  //   注文アプリ内で登録した方の 性別・種別・生年・誕生日 が捨てられていた。
  //   その結果 ①会員IDが常に WN-G-xxxx（性別・学生/シニア/観光客の区別が消える）
  //   ②シニア自動判定が効かない ③誕生日メールが一生届かない、という状態だった。
  if (data.isMember && data.memberContact) {
    ensureMember(ss, {
      email:      data.memberContact,
      name:       data.memberName || "",
      birthYear:  data.memberBirthYear  || "",
      birthMonth: data.memberBirthMonth || "",
      birthDay:   data.memberBirthDay   || "",
      language:   data.memberLanguage || "",
      gender:     data.memberGender   || "",
      memberType: data.memberType     || "G",
    });
  }

  // ── 同席会員へのスタンプ自動付与 ─────────────────────────────
  const coseatResults = [];
  const coseatMembers = data.coseatMembers || [];
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");

  // 当日注文のラーメン杯数を集計（テーブル単位）
  // ※クーポン使用時は半額/無料の1杯を除外（その丼はポイント対象外のため）
  const ramenQty = (data.items || []).reduce((sum, item) => {
    const name = (item.namePL || item.nameEN || "").toLowerCase();
    const isR = name.includes("ramen") || name.includes("tantanmen") ||
           name.includes("tonkotsu") || name.includes("miso") ||
           name.includes("udon") || name.includes("soba") ||
           name.includes("tsukemen") || name.includes("assari") ||
           name.includes("paitan") || name.includes("shoyu") ||
           name.includes("shio") || name.includes("kimchi") ||
           name.includes("seafood") || name.includes("curry") ||
           name.includes("zaru") || name.includes("wabi navi") ||
           name.includes("wege") || name.includes("salad ramen") ||
           name.includes("inari udon");
    return sum + (isR ? (Number(item.qty) || 1) : 0);
  }, 0);
  const orderHasRamen = (ramenQty - (couponUsed ? 1 : 0)) > 0;

  coseatMembers.forEach(email => {
    email = email.trim().toLowerCase();
    if (!email) return;

    const member = findMember(email);
    if (!member) {
      coseatResults.push({ email, status:"notFound" });
      return;
    }

    // 失効チェック
    if (checkMemberExpiry(member)) {
      updateMember(ss, email, { points:0, rank:"Sakura", _source:"失効リセット / Expired (order)" });
      coseatResults.push({ email, status:"expired" });
      return;
    }

    // 本日スタンプ済みチェック
    const history = member.history || [];
    if (history.includes(today)) {
      coseatResults.push({ email, status:"alreadyStamped" });
      return;
    }

    // ラーメンなし
    if (!orderHasRamen) {
      coseatResults.push({ email, status:"noRamen" });
      return;
    }

    // ポイント付与（⏰ハッピーアワー中は加算）
    const hh     = happyHourStatus();
    const oldPt  = member.points || 0;
    const newPt  = oldPt + 10 + (hh.active ? hh.bonus : 0);
    const bonus  = awardMilestoneCoupons(email, member.name, oldPt, newPt);
    const newRank = getRank(newPt);
    const rankUp  = getRank(oldPt).name !== newRank.name;

    history.push(today);
    updateMember(ss, email, { points:newPt, history, lastVisit:today, rank:newRank.name,
      _source: "Pieczątka z zamówienia — dosiadka / Order stamp (co-seat)" +
               (hh.active ? " ⏰HH+" + hh.bonus : "") });
    coseatResults.push({
      email, status:"ok", oldPt, newPt,
      rank:newRank, rankUp, bonus
    });
  });

  return jsonResponse({ status:"ok", orderId, coseatResults });
}

// ════════════════════════════════════════════════════════════════════
//  🎌 食で学ぶ日本語 — 講座＋クイズ（v6.0）
//    月2回、店主による3分の日本語講座（YouTube公開）。
//    クイズ正解で自動的に +5pt。1会員1回号1回まで。
//    正解者は抽選対象として応募シートに記録される。
// ════════════════════════════════════════════════════════════════════
const SHEET_LESSONS = "日本語講座 / Lekcje";
const SHEET_QUIZLOG = "クイズ応募 / Quiz";

function setupLessonSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let sh = ss.getSheetByName(SHEET_LESSONS);
  if (sh) {
    const r = ui.alert("講座シートは既に存在します",
      "作り直しますか？（現在の内容は消えます）", ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet(SHEET_LESSONS);
  const headers = [
    "回号 / Nr",                    // A
    "公開日 / Data (yyyy-mm-dd)",   // B
    "タイトルPL",                   // C
    "タイトルJP",                   // D
    "YouTube ID（v=の後ろ）",        // E
    "フレーズJP / 日本語フレーズ",    // F
    "よみ / Romaji",                // G
    "意味PL / Znaczenie",           // H
    "クイズ問題PL",                  // I
    "選択肢1",                      // J
    "選択肢2",                      // K
    "選択肢3",                      // L
    "選択肢4（空欄可）",             // M
    "正解番号 1-4",                 // N
    "応募締切 / Do (yyyy-mm-dd)",   // O
    "特典 / Nagroda",               // P
  ];
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold")
    .setBackground("#8e44ad").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setFontSize(9);
  [70, 140, 220, 180, 190, 200, 160, 220, 260, 150, 150, 150, 150, 100, 140, 180]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  sh.appendRow([1, "2026-08-01", "Lekcja 1: Zamawianie ramenu", "第1回 ラーメンの注文",
    "dQw4w9WgXcQ", "おまかせします", "omakase shimasu",
    "Zdaję się na Pana wybór / I'll leave it to you",
    "Co znaczy「おまかせします」?",
    "Poproszę rachunek", "Zdaję się na Pana wybór", "Jest bardzo ostre", "", 2,
    "2026-08-31", "Kupon 10 zł / 10zł割引券"]);
  sh.getRange(2, 1, 1, 16).setFontColor("#999999").setFontStyle("italic");

  // 応募ログシート
  let q = ss.getSheetByName(SHEET_QUIZLOG);
  if (!q) {
    q = ss.insertSheet(SHEET_QUIZLOG);
    q.appendRow(["Data / Czas", "Lekcja / Nr", "ID członka", "Imię / Name",
                 "Kontakt", "Odpowiedź", "Wynik / Result", "Punkty", "Wygrana / Prize"]);
    q.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#8e44ad")
      .setFontColor("#ffffff").setHorizontalAlignment("center").setFontSize(9);
    [150, 60, 130, 140, 200, 70, 90, 70, 110].forEach((w, i) => q.setColumnWidth(i + 1, w));
    q.setFrozenRows(1);
  }
  try { CacheService.getScriptCache().remove("lessons_v1"); } catch (e) {}
  ui.alert("✅ 講座シートを作成しました",
    "・1行＝1回分の講座です（記入例は削除してOK）\n" +
    "・E列のYouTube IDは、動画URLの v= の後ろの文字列だけ\n" +
    "・N列に正解番号（1〜4）\n" +
    "・公開日が未来の回は、まだサイトに出ません\n\n" +
    "クイズ正解者には自動で+5ptが付き、応募シートに記録されます。",
    ui.ButtonSet.OK);
}

// ── 講座機能のON/OFF設定（v6.4）────────────────────────────────
//   LESSON_MODE : "on"（公開）/ "soon"（準備中表示）/ "off"（完全非表示）
//   QUIZ_POINTS : "on"（正解で+5pt）/ "off"（採点のみ・ポイントなし）
//   ※初期値は「準備中・ポイントなし」。メニューからいつでも切替可能。
function getLessonSetting(key, def) {
  try { return PropertiesService.getScriptProperties().getProperty(key) || def; }
  catch (e) { return def; }
}
function lessonMode()  { return getLessonSetting("LESSON_MODE", "soon"); }
function quizPointsOn() { return getLessonSetting("QUIZ_POINTS", "off") === "on"; }

// メニュー：講座の公開状態とポイント付与を切り替える
function toggleLessonSettings() {
  const ui = SpreadsheetApp.getUi();
  const cur  = lessonMode();
  const curP = quizPointsOn() ? "ON" : "OFF";
  const label = { on:"1 = ON (widoczne / 公開中)", soon:"2 = WKRÓTCE (coming soon / 準備中)", off:"3 = OFF (ukryte / 非表示)" };

  const r1 = ui.prompt("🎌 Lekcje japońskie / Japanese lessons",
    "Stan obecny / Current: " + (label[cur] || cur) + "\n\n" +
    "Wybierz nowy stan — wpisz 1, 2 lub 3:\n" +
    "  1 … ON — lekcje widoczne w aplikacji i na stronie\n" +
    "  2 … WKRÓTCE — pokazuje „coming soon”\n" +
    "  3 … OFF — całkowicie ukryte\n" +
    "（1=公開 / 2=準備中 / 3=非表示）",
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const v = r1.getResponseText().trim();
  const map = { "1":"on", "2":"soon", "3":"off" };
  if (!map[v]) { ui.alert("⚠️ Wpisz 1, 2 albo 3 / Enter 1, 2 or 3"); return; }

  const r2 = ui.alert("🎁 Punkty za quiz / Quiz points",
    "Czy poprawna odpowiedź ma dawać +5 pkt?\n" +
    "Should a correct answer award +5 points?\n" +
    "（クイズ正解で+5ポイントを付与しますか？）\n\n" +
    "Teraz / Now: " + curP,
    ui.ButtonSet.YES_NO);

  const props = PropertiesService.getScriptProperties();
  props.setProperty("LESSON_MODE", map[v]);
  props.setProperty("QUIZ_POINTS", r2 === ui.Button.YES ? "on" : "off");
  try { CacheService.getScriptCache().remove("lessons_v1"); } catch (e) {}

  ui.alert("✅ Zapisano / Saved",
    "Lekcje / Lessons: " + (label[map[v]] || map[v]) + "\n" +
    "Punkty / Points: " + (r2 === ui.Button.YES ? "ON (+5pt)" : "OFF") + "\n\n" +
    "Zmiana widoczna w ciągu ~5 minut (cache).\n反映は最大5分（キャッシュ）です。",
    ui.ButtonSet.OK);
}

// GET action=getLessons — 公開済みの講座一覧（正解番号は返さない）
function getLessonsPublic() {
  const mode   = lessonMode();
  const qPoints = quizPointsOn();
  // 非公開状態なら講座データ自体を返さない（ページ側は「準備中」を表示）
  if (mode !== "on") return { status: "ok", mode, quizPoints: qPoints, lessons: [] };

  try {
    const cached = CacheService.getScriptCache().get("lessons_v1");
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LESSONS);
  if (!sh) return { status: "noSheet", mode, quizPoints: qPoints, lessons: [] };

  const rows  = sh.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const list  = [];

  for (let i = 1; i < rows.length; i++) {
    const r  = rows[i];
    const no = Number(r[0]);
    if (!no) continue;
    const pub = fmtPromoDate(r[1]);
    if (pub && today < pub) continue;             // 未公開回は出さない

    const opts = [r[9], r[10], r[11], r[12]]
      .map(x => (x || "").toString().trim()).filter(Boolean);
    const deadline = fmtPromoDate(r[14]);
    list.push({
      no, date: pub,
      titlePL: (r[2] || "").toString(), titleJP: (r[3] || "").toString(),
      youtube: (r[4] || "").toString().trim(),
      phraseJP: (r[5] || "").toString(), romaji: (r[6] || "").toString(),
      meaning: (r[7] || "").toString(),
      quiz: (r[8] || "").toString(), options: opts,
      deadline, prize: (r[15] || "").toString(),
      open: !deadline || today <= deadline        // 応募受付中か
      // ※正解番号(N列)は意図的に返さない（ページのソースから答えが見えないように）
    });
  }
  list.sort((a, b) => b.no - a.no);               // 新しい回が先頭

  const result = { status: "ok", mode, quizPoints: qPoints, lessons: list, today };
  try { CacheService.getScriptCache().put("lessons_v1", JSON.stringify(result), 300); } catch (e) {}
  return result;
}

// POST action=submitQuiz {memberId, lessonNo, answer}
function handleQuizSubmit(data) {
  if (lessonMode() !== "on") {
    return jsonResponse({ status: "closed",
      message: "Quiz jest obecnie niedostępny / Quiz is not available now（講座は現在準備中です）" });
  }
  const memberId = (data.memberId || "").toString().trim();
  const lessonNo = Number(data.lessonNo);
  const answer   = Number(data.answer);
  if (!memberId || !lessonNo || !answer) {
    return jsonResponse({ status: "error", message: "Brak danych / Missing data" });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LESSONS);
  if (!sh) return jsonResponse({ status: "error", message: "講座シートがありません" });

  // 会員特定（会員ID・メール・電話のいずれでも可）
  const member = findMember(memberId);
  if (!member) {
    return jsonResponse({ status: "notMember",
      message: "会員が見つかりません。会員登録がお済みか、IDをご確認ください。\nNie znaleziono członka — sprawdź swój ID." });
  }

  // 該当回を探す
  const rows = sh.getDataRange().getValues();
  let row = null;
  for (let i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === lessonNo) { row = rows[i]; break; }
  }
  if (!row) return jsonResponse({ status: "error", message: "その回が見つかりません" });

  const today    = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const pub      = fmtPromoDate(row[1]);
  const deadline = fmtPromoDate(row[14]);
  if (pub && today < pub)      return jsonResponse({ status: "error", message: "この回はまだ公開されていません" });
  if (deadline && today > deadline) {
    return jsonResponse({ status: "closed",
      message: "この回の応募は締め切りました / Termin minął（" + deadline + "）" });
  }

  // 重複応募チェック（1会員1回号1回）
  let q = ss.getSheetByName(SHEET_QUIZLOG);
  if (!q) {
    q = ss.insertSheet(SHEET_QUIZLOG);
    q.appendRow(["Data / Czas", "Lekcja / Nr", "ID członka", "Imię / Name",
                 "Kontakt", "Odpowiedź", "Wynik / Result", "Punkty", "Wygrana / Prize"]);
  }
  const qRows = q.getDataRange().getValues();
  const myId  = (member.memberId || "").toString().trim().toLowerCase();
  for (let i = 1; i < qRows.length; i++) {
    if (Number(qRows[i][1]) === lessonNo &&
        (qRows[i][2] || "").toString().trim().toLowerCase() === myId) {
      return jsonResponse({ status: "already",
        message: "この回はすでに応募済みです / Już wysłano odpowiedź（" + (member.name || "") + "）" });
    }
  }

  const correct = Number(row[13]);
  const isRight = (answer === correct);
  const stamp   = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd HH:mm");
  let awarded   = 0;
  let newPt     = member.points || 0;
  let rankUp    = false;

  const givePoints = quizPointsOn();
  if (isRight && givePoints) {
    const oldPt = member.points || 0;
    newPt   = oldPt + 5;
    awarded = 5;
    rankUp  = getRank(oldPt).name !== getRank(newPt).name;
    try { awardMilestoneCoupons(member.email, member.name, oldPt, newPt); }
    catch (err) { Logger.log("quiz coupon: " + err); }
    // クイズ分はポイントのみ加算（来店履歴＝杯数には数えない）
    updateMember(ss, member.email, {
      points: newPt, rank: getRank(newPt).name,
      _source: "Quiz — poprawna odpowiedź / Quiz correct (lekcja " + lessonNo + ")"
    });
  }

  q.appendRow([stamp, lessonNo, member.memberId || "", member.name || "",
               member.email || "", answer, isRight ? "⭕ Poprawnie" : "❌ Błędnie",
               awarded, ""]);

  return jsonResponse({
    status: "ok", correct: isRight, awarded, newPt, rankUp,
    quizPoints: givePoints,
    rank: getRank(newPt),
    memberName: member.name || "",
    prize: (row[15] || "").toString()
  });
}

// ════════════════════════════════════════════════════════════════════
//  スタンプ処理
// ════════════════════════════════════════════════════════════════════

// ── QRスキャン付与（係員用・v5.1） ──────────────────────────────
//  会員カードのQR（会員ID）をスタッフがスキャンして呼び出す。
//  流れ：①会員IDで会員特定 → ②失効/1日1回チェック →
//        ③当日注文に本人メールが紐付いていれば自動付与 →
//        ④紐付きなし（同席者の付与忘れ等）は needConfirm を返し、
//          スタッフが確認後 force:true で再送 → 同席救済として付与
//  ※通常スタンプと同じ防壁（1日1回・失効）はforce時も維持される
function handleStaffQrStamp(data) {
  const auth1 = verifyStaffPassword(data.password);
  if (auth1 !== "ok") return jsonResponse({ status:"unauthorized", message: STAFF_AUTH_MSG[auth1] });
  const memberId = (data.memberId || "").toString().trim();
  if (!memberId) return jsonResponse({ status:"error", message:"会員IDが必要です / Member ID required" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const member = findMember(memberId);  // findMemberはA列（会員ID）でも検索できる
  if (!member) {
    return jsonResponse({ status:"notFound", message:"この会員IDは登録されていません / Nie znaleziono: " + memberId });
  }
  const email = (member.email || "").toString().trim().toLowerCase();

  // 失効チェック（8ヶ月来店なし）
  if (checkMemberExpiry(member)) {
    updateMember(ss, email, { points:0, rank:"Sakura", _source:"Punkty wygasły / Expired (QR staff)" });
    return jsonResponse({
      status:"expired",
      message:"ポイント有効期限切れ（8ヶ月来店なし）。0ptから再スタートしました。\nPunkty wygasły — restart od 0pt."
    });
  }

  // 1日1回チェック
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const history = member.history || [];
  if (history.includes(today)) {
    return jsonResponse({
      status:"alreadyStamped",
      message:"本日はすでにスタンプ済みです / Pieczątka już dziś przyznana（" + (member.name||"") + "）"
    });
  }

  // 当日注文照合（本人メールが注文に紐付いているか）
  const hasRamen = checkTodayRamen(ss, email, today);
  if (!hasRamen && !data.force) {
    return jsonResponse({
      status:"needConfirm",
      member:{ memberId: member.memberId, name: member.name, rank: member.rank, points: member.points }
    });
  }

  // ポイント付与（通常スタンプと同一ロジック・⏰ハッピーアワー加算）
  const hh     = happyHourStatus();
  const oldPt  = member.points || 0;
  const newPt  = oldPt + 10 + (hh.active ? hh.bonus : 0);
  const bonus  = awardMilestoneCoupons(email, member.name, oldPt, newPt);
  const newRank = getRank(newPt);
  const rankUp  = getRank(oldPt).name !== newRank.name;

  history.push(today);
  updateMember(ss, email, {
    points: newPt, history, lastVisit: today, rank: newRank.name,
    _source: (hh.active ? "⏰HH+" + hh.bonus + " " : "") +
             (hasRamen ? "Skan QR przez obsługę / QR scan (staff)"
                      : "Dosiadka — zatwierdzone przez obsługę / Co-seat (staff approved)")
  });

  return jsonResponse({
    status:"ok", oldPt, newPt, rank:newRank, rankUp, bonus,
    happyHour: hh.active ? hh.bonus : 0,
    coseat: !hasRamen,
    member:{ memberId: member.memberId, name: member.name }
  });
}

function handleStamp(data) {
  const email = (data.email || "").trim().toLowerCase();
  if (!email) return jsonResponse({ status:"error", message:"メールアドレスが必要です / Email required" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 会員確認
  const member = findMember(email);
  if (!member) {
    return jsonResponse({
      status: "notMember",
      message: "このメールアドレスは登録されていません。\nThis email is not registered.\n会員登録はアプリまたはレジにてどうぞ。"
    });
  }

  // 失効チェック（8ヶ月来店なし）
  if (checkMemberExpiry(member)) {
    const ss2 = SpreadsheetApp.getActiveSpreadsheet();
    // 履歴は保持・ポイントとランクのみリセット
    updateMember(ss2, email, { points:0, rank:"Sakura", _source:"失効リセット / Expired (stamp)" });
    return jsonResponse({
      status: "expired",
      message: "ポイントの有効期限が切れました（8ヶ月間のご来店なし）。\n0ptから再スタートします！\nPunkty wygasły (8 miesięcy bez wizyty). Zaczynamy od nowa!\nYour points have expired. Fresh start from 0pt!",
      newMember: { ...member, points:0, rank:"Sakura" }
    });
  }

  // 本日スタンプ済みチェック
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const history = member.history || [];
  if (history.includes(today)) {
    return jsonResponse({
      status: "alreadyStamped",
      message: "本日はすでにスタンプが押されています！\nYou already received a stamp today!\nNa dzisiaj już otrzymałeś/aś pieczątkę!",
      member
    });
  }

  // 当日注文にラーメンが含まれるか照合
  const hasRamen = checkTodayRamen(ss, email, today);
  if (!hasRamen) {
    return jsonResponse({
      status: "noRamen",
      message: "本日のご注文にラーメンが含まれていません。\nW dzisiejszym zamówieniu nie ma ramenu.\nNo ramen found in today's order.\n毎回のご来店時、ラーメンご注文のお客様に1回限り適用されます。"
    });
  }

  // ポイント付与（⏰ハッピーアワー中は加算）
  const hh     = happyHourStatus();
  const oldPt  = member.points || 0;
  const newPt  = oldPt + 10 + (hh.active ? hh.bonus : 0);
  const bonus  = awardMilestoneCoupons(email, member.name, oldPt, newPt);
  const oldRank = getRank(oldPt);
  const newRank = getRank(newPt);
  const rankUp  = oldRank.name !== newRank.name;

  history.push(today);
  updateMember(ss, email, { points: newPt, history, lastVisit: today, rank: newRank.name,
    _source: "Pieczątka QR / Stamp QR" + (hh.active ? " ⏰HH+" + hh.bonus : "") });

  return jsonResponse({
    status:   "ok",
    oldPt, newPt,
    rank:     newRank,
    rankUp,
    bonus,
    message:  `+10pt！　累計 ${newPt}pt`,
    member:   { ...member, points: newPt, rank: newRank.name, history }
  });
}

// ════════════════════════════════════════════════════════════════════
//  会員登録処理
// ════════════════════════════════════════════════════════════════════
function handleRegister(data) {
  const email = (data.email || "").trim().toLowerCase();
  const name  = (data.name  || "").trim();
  if (!email || !name) {
    return jsonResponse({ status:"error", message:"名前とメールアドレスが必要です / Name and email required" });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = findMember(email);
  if (existing) {
    return jsonResponse({ status:"alreadyExists", message:"すでに登録されています / Already registered", member: existing });
  }

  // シニア自動判定（サーバー側でも確認）
  let memberType = data.memberType || "G";
  if (data.birthYear) {
    const age = new Date().getFullYear() - parseInt(data.birthYear);
    if (age >= 65) memberType = "SR";
  }

  const newMember = {
    email,
    name,
    birthYear:   data.birthYear   || "",
    birthMonth:  data.birthMonth  || "",
    birthDay:    data.birthDay    || "",
    language:    data.language    || "",
    gender:      data.gender      || "",
    memberType,
    points:      0,
    rank:        "Sakura",
    registeredAt: Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd"),
    lastVisit:   "",
    history:     [],
  };

  saveMember(ss, newMember);

  // saveMember が採番した会員IDを取り込む（メールに記載するため）
  try {
    const saved = findMember(email);
    if (saved && saved.memberId) newMember.memberId = saved.memberId;
  } catch (e) {}

  // 登録完了メール（WhatsApp参加リンク入り）。失敗しても登録自体は成功扱い。
  try { sendWelcomeMail(newMember); }
  catch (e) { Logger.log("welcome mail: " + e); }

  // 画面に「WhatsAppに参加」ボタンを出すためリンクも返す
  return jsonResponse({ status:"ok", member: newMember, whatsapp: getWhatsAppUrl() });
}

// ════════════════════════════════════════════════════════════════════
//  会員カード取得
// ════════════════════════════════════════════════════════════════════
function handleGetCard(data) {
  const email = (data.email || "").trim().toLowerCase();
  if (!email) return jsonResponse({ status:"error", message:"Email required" });

  const member = findMember(email);
  if (!member) return jsonResponse({ status:"notFound" });

  const rank = getRank(member.points || 0);
  const coupons = getUnusedCoupons(email);
  return jsonResponse({ status:"ok", member: { ...member, rank: rank.name, rankIcon: rank.icon, coupons } });
}

// ════════════════════════════════════════════════════════════════════
//  退会処理（GDPR対応・全データ削除）
// ════════════════════════════════════════════════════════════════════
function handleWithdraw(data) {
  const email    = (data.email    || "").trim().toLowerCase();
  const memberId = (data.memberId || "").trim().toLowerCase();
  if (!email) return jsonResponse({ status:"error", message:"Email required" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getMemberSheet(ss);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const rowEmail = (rows[i][1] || "").toString().trim().toLowerCase();
    if (rowEmail === email) {
      // v5.6：メールだけで退会（＝全データ消去）できた穴を修正。
      // 本人のカード画面にしか表示されない会員IDの一致を必須にする。
      // （会員IDが空の旧データ行のみ、従来通りメールだけで許可）
      const rowId = (rows[i][0] || "").toString().trim().toLowerCase();
      if (rowId && rowId !== memberId) {
        return jsonResponse({ status:"error",
          message:"本人確認に失敗しました。会員カード画面から退会操作を行ってください。\nWeryfikacja nie powiodła się — użyj ekranu karty członka." });
      }
      // 行を削除（GDPRに基づく完全消去）
      sh.deleteRow(i + 1);

      // ポイント履歴シートからも該当会員の記録を完全削除（GDPR）
      const logSh = ss.getSheetByName(SHEET_POINTLOG);
      if (logSh) {
        const logData = logSh.getDataRange().getValues();
        for (let j = logData.length - 1; j >= 1; j--) {
          if ((logData[j][2] || "").toString().trim().toLowerCase() === email) {
            logSh.deleteRow(j + 1);
          }
        }
      }

      // 退会確認メールを送信
      sendWithdrawConfirmMail(email, rows[i][2] || "");

      Logger.log("Withdrawn: " + email);
      return jsonResponse({ status:"ok", message:"退会処理完了 / Membership withdrawn" });
    }
  }
  return jsonResponse({ status:"notFound", message:"会員が見つかりません / Member not found" });
}

// ── 退会確認メール ─────────────────────────────────────────────────
function sendWithdrawConfirmMail(email, name) {
  const subject = "Wabi Navi — Potwierdzenie rezygnacji / Membership Withdrawal Confirmation / 退会完了のご連絡";
  const body = `
${name ? name + "様 / Drogi/a " + name + "," : "Drogi/a Członku/Członkini,"}

━━━━━━━━━━━━━━━━━━━━━━━━━━
🌸 Dziękujemy za lojalność!
   Thank you for your loyalty!
   ご愛顧ありがとうございました。
━━━━━━━━━━━━━━━━━━━━━━━━━━

Twoje członkostwo w Wabi Navi Members Club zostało anulowane.
Your Wabi Navi Members Club membership has been cancelled.
Wabi Navi Members Clubの退会手続きが完了しました。

Wszystkie Twoje dane osobowe zostały trwale usunięte z naszego systemu.
All your personal data has been permanently deleted from our system.
保存されていた個人データはすべて消去されました。

Zapraszamy ponownie w każdej chwili!
We look forward to welcoming you back anytime!
またのご来店を心よりお待ちしております。

📍 Wabi Navi — ul. Małe Garbary 5A, 87-100 Toruń
🌐 ${SITE_URL}

━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI
━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  try {
    GmailApp.sendEmail(email, subject, body);
  } catch(e) {
    Logger.log("Withdraw mail error: " + e.toString());
  }
}

// ════════════════════════════════════════════════════════════════════
//  係員：手動ポイント追加
// ════════════════════════════════════════════════════════════════════
function handleStaffAddPoint(data) {
  const auth2 = verifyStaffPassword(data.password);
  if (auth2 !== "ok") return jsonResponse({ status:"error", message: STAFF_AUTH_MSG[auth2] });

  const email = (data.email || "").trim().toLowerCase();
  const member = findMember(email);
  if (!member) return jsonResponse({ status:"notFound" });

  // ポイント数：10（通常）/ 20（ダブル・2杯）/ 5（🎌クイズ正解）のみ許可
  const raw = Number(data.points);
  const addPts = (raw === 20) ? 20 : (raw === 5) ? 5 : 10;
  const isQuiz = (addPts === 5);

  const oldPt = member.points || 0;
  const newPt = oldPt + addPts;
  const bonus  = awardMilestoneCoupons(email, member.name, oldPt, newPt);
  const newRank = getRank(newPt);

  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const history = member.history || [];
  const upd = { points: newPt, lastVisit: today, rank: newRank.name,
    _source: isQuiz      ? "Panel obsługi — quiz +5 / Staff panel quiz" :
             addPts === 20 ? "Panel obsługi — podwójny / Staff panel x2"
                           : "Panel obsługi / Staff panel" };
  if (!isQuiz) {
    history.push(today);
    if (addPts === 20) history.push(today);  // ダブル＝2杯分の来店履歴
    upd.history = history;                    // クイズ+5は杯数に数えない
  }

  updateMember(SpreadsheetApp.getActiveSpreadsheet(), email, upd);

  return jsonResponse({
    status: "ok",
    oldPt, newPt,
    rank: newRank,
    bonus,
    message: `+${addPts}pt追加（係員操作${addPts===20?"・ダブル":isQuiz?"・クイズ":""}）　累計 ${newPt}pt`
  });
}

// ── 係員：手動ポイント取消（スマホ係員パネル用・v5.5） ──────────
//  -10pt：当日付与分がある場合のみ。来店履歴からも当日分を1件削除
//  -5pt ：クイズ分の誤付与用。ポイントのみ減算
function handleStaffRemovePoint(data) {
  const auth3 = verifyStaffPassword(data.password);
  if (auth3 !== "ok") return jsonResponse({ status:"error", message: STAFF_AUTH_MSG[auth3] });
  const email = (data.email || "").trim().toLowerCase();
  const member = findMember(email);
  if (!member) return jsonResponse({ status:"notFound", message:"未登録の会員です / Nie znaleziono" });

  const amt = (Number(data.points) === 5) ? 5 : 10;
  const today      = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const history    = member.history || [];
  const todayCount = history.filter(d => d === today).length;
  const oldPt      = member.points || 0;

  if (oldPt < amt) {
    return jsonResponse({ status:"error",
      message:"現在のポイント（" + oldPt + "pt）が取消額より少ないため取消できません" });
  }
  if (amt === 10 && todayCount === 0) {
    return jsonResponse({ status:"error",
      message:"本日付与された🍜ポイントがありません（-10ptの取消は当日のみ可能）" });
  }

  const newPt    = Math.max(0, oldPt - amt);
  const newRank  = getRank(newPt);
  const rankDown = getRank(oldPt).name !== newRank.name;

  const upd = { points:newPt, rank:newRank.name,
    _source: amt === 10 ? "Panel obsługi — cofnięcie -10 / Staff panel undo"
                        : "Panel obsługi — cofnięcie -5 / Staff panel undo quiz" };
  if (amt === 10) {
    const newHistory = [...history];
    const idx = newHistory.lastIndexOf(today);
    if (idx !== -1) newHistory.splice(idx, 1);
    upd.history   = newHistory;
    upd.lastVisit = newHistory.length ? newHistory[newHistory.length - 1] : "";
  }
  updateMember(SpreadsheetApp.getActiveSpreadsheet(), email, upd);

  return jsonResponse({
    status:"ok", oldPt, newPt, rank:newRank, rankDown,
    message:"-" + amt + "pt取消しました　累計 " + newPt + "pt" + (rankDown ? "（ランクダウン）" : "")
  });
}

// ════════════════════════════════════════════════════════════════════
//  当日のラーメン注文を照合
// ════════════════════════════════════════════════════════════════════
function checkTodayRamen(ss, email, today) {
  const sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) return false;

  const data   = sh.getDataRange().getValues();
  const todayD = new Date(today);
  let ramenQty = 0, couponFlag = false;

  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const rowDate   = new Date(row[1]);
    const rowEmail  = (row[14] || "").toString().trim().toLowerCase();
    const itemName  = (row[5]  || "").toString().toLowerCase();
    const rowUwagi  = (row[11] || "").toString();

    const sameDay = rowDate.toDateString() === todayD.toDateString();
    const sameUser = rowEmail === email;
    const isRamen  = itemName.includes("ramen") || itemName.includes("tantanmen") ||
                     itemName.includes("tonkotsu") || itemName.includes("miso") ||
                     itemName.includes("udon") || itemName.includes("soba") ||
                     itemName.includes("tsukemen") || itemName.includes("assari") ||
                     itemName.includes("paitan") || itemName.includes("shoyu") ||
                     itemName.includes("shio") || itemName.includes("kimchi") ||
                     itemName.includes("seafood") || itemName.includes("curry") ||
                     itemName.includes("zaru") || itemName.includes("wabi navi") ||
                     itemName.includes("wege") || itemName.includes("salad ramen") ||
                     itemName.includes("inari udon");

    if (sameDay && sameUser) {
      if (isRamen) ramenQty += Number(row[7]) || 1;
      if (rowUwagi.indexOf("🎫 KUPON") >= 0) couponFlag = true;
    }
  }
  // ルール：クーポン（半額・無料とも）を使用した来店はスタンプ対象外
  if (couponFlag) return false;
  return ramenQty > 0;
}

// ════════════════════════════════════════════════════════════════════
//  会員シート操作
// ════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════
//  会員ID生成
//  形式：WN-[属性コード]-[通し番号4桁]
//  属性コード：
//    性別：M=男性 F=女性 （未回答は省略）
//    種別：G=一般 S=学生 SR=シニア T=観光客
//  例：WN-MG-0001 / WN-FS-0023 / WN-SR-0015 / WN-T-0003
// ════════════════════════════════════════════════════════════════════

function generateMemberId(gender, memberType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getMemberSheet(ss);

  // 通し番号：既存会員数+1（全種別共通）
  const lastRow = sh.getLastRow();
  const seq     = String(Math.max(lastRow, 1)).padStart(4, '0');

  // 属性コード組み立て
  let code = '';

  // 性別コード（未回答は省略）
  if (gender === 'M' || gender === 'male'   || gender === 'mężczyzna') code += 'M';
  if (gender === 'F' || gender === 'female' || gender === 'kobieta')   code += 'F';

  // 種別コード
  if      (memberType === 'tourist' || memberType === 'T')  code += 'T';
  else if (memberType === 'senior'  || memberType === 'SR') code += 'SR';
  else if (memberType === 'student' || memberType === 'S')  code += 'S';
  else                                                       code += 'G';

  return 'WN-' + code + '-' + seq;
}

function getMemberSheet(ss) {
  let sh = ss.getSheetByName(SHEET_MEMBERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MEMBERS);
    const headers = [
      "🪪 会員ID / Member ID",           // A
      "Email / 連絡先",                  // B
      "氏名 / Name",                     // C
      "登録日 / Registered",             // D
      "生年 / Birth year",               // E ★新規
      "誕生月 / Birth month",            // F
      "誕生日 / Birth day",              // G ★新規（任意）
      "言語 / Language",                 // H
      "性別 / Gender",                   // I
      "種別 / Member type",              // J
      "累計ポイント / Total points",     // K
      "ランク / Rank",                   // L
      "🎖️ ランクアイコン",               // M
      "次のランクまで / To next rank",   // N
      "最終来店日 / Last visit",         // O
      "来店回数 / Visit count",          // P
      "来店履歴 / Visit history"         // Q
    ];
    sh.appendRow(headers);
    const hdr = sh.getRange(1,1,1,headers.length);
    hdr.setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e")
       .setHorizontalAlignment("center");
    sh.setColumnWidth(1, 140);  // A 会員ID
    sh.setColumnWidth(2, 200);  // B Email
    sh.setColumnWidth(3, 150);  // C Name
    sh.setColumnWidth(4, 110);  // D Registered
    sh.setColumnWidth(5, 80);   // E Birth year
    sh.setColumnWidth(6, 90);   // F Birth month
    sh.setColumnWidth(7, 75);   // G Birth day
    sh.setColumnWidth(8, 75);   // H Language
    sh.setColumnWidth(9, 90);   // I Gender
    sh.setColumnWidth(10, 110); // J Member type
    sh.setColumnWidth(11, 130); // K Total points
    sh.setColumnWidth(12, 100); // L Rank name
    sh.setColumnWidth(13, 80);  // M Rank icon
    sh.setColumnWidth(14, 160); // N To next rank
    sh.setColumnWidth(15, 120); // O Last visit
    sh.setColumnWidth(16, 100); // P Visit count
    sh.setColumnWidth(17, 300); // Q Visit history
    sh.setFrozenRows(1);
  }
  return sh;
}

function findMember(emailOrPhone) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = getMemberSheet(ss);
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // v5.4修正：A列=会員ID / B列=連絡先（登録時のメール or 電話）で照合。
    // 旧版はC列（氏名）とも比較していた（誤動作の元）ため除去。
    // 電話番号は末尾9桁照合で +48/スペース等の表記ゆれを吸収。
    if (memberRowMatches(data[i][0], data[i][1], emailOrPhone)) {
      // v7.0修正：来店履歴は Q列（index 16）。旧版は index 8（I列＝性別）を
      // 読んでいたため、①1日1回のスタンプ制限が効かない ②来店回数が常に2
      // ③−10ptの取消が常に拒否される、という3つの不具合が出ていた。
      // 書き込み側（updateMember）は元から Q列(17) で正しい。
      const history = data[i][16] ? data[i][16].toString().split(",").filter(Boolean) : [];
      return {
        memberId:    data[i][0],
        email:       data[i][1],
        name:        data[i][2],
        registeredAt:data[i][3],
        birthYear:   data[i][4],
        birthMonth:  data[i][5],
        birthDay:    data[i][6],
        language:    data[i][7],
        gender:      data[i][8],
        memberType:  data[i][9],
        points:      Number(data[i][10]) || 0,
        rank:        data[i][11],
        lastVisit:   data[i][14],
        history,
        row: i+1
      };
    }
  }
  return null;
}

function saveMember(ss, m) {
  const sh = getMemberSheet(ss);
  const rankInfo  = getRank(m.points || 0);
  const nextMs    = getNextMilestone(m.points || 0);
  const toNext    = nextMs !== null ? (nextMs - (m.points||0)) + "pt → " + getNextRankName(m.points||0) : "🎖️ MAX (Fuji)";
  const memberId  = m.memberId || generateMemberId(m.gender || '', m.memberType || 'G');
  const genderLabel = m.gender === 'M' ? '男性 / Male' :
                      m.gender === 'F' ? '女性 / Female' : '未回答 / —';
  const typeLabel   = m.memberType === 'S'  ? '学生 / Student' :
                      m.memberType === 'SR' ? 'シニア / Senior' :
                      m.memberType === 'T'  ? '観光客 / Tourist' : '一般 / General';
  sh.appendRow([
    memberId,                    // A 会員ID
    m.email,                     // B Email
    m.name,                      // C Name
    m.registeredAt,              // D Registered
    m.birthYear   || "",         // E 生年
    m.birthMonth  || "",         // F 誕生月
    m.birthDay    || "",         // G 誕生日（任意）
    m.language    || "",         // H Language
    genderLabel,                 // I 性別
    typeLabel,                   // J 種別
    m.points,                    // K Total points
    m.rank,                      // L Rank
    rankInfo.icon,               // M Rank icon
    toNext,                      // N To next rank
    m.lastVisit   || "",         // O Last visit
    (m.history||[]).length,      // P Visit count
    (m.history||[]).join(",")    // Q Visit history
  ]);
}

function updateMember(ss, email, updates) {
  const sh   = getMemberSheet(ss);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // v5.4修正：旧版はA列（会員ID）とメールを比較していたため一度も一致せず、
    // ポイント更新が全会員で静かに失敗していた。ID/連絡先の両対応照合に変更。
    if (memberRowMatches(data[i][0], data[i][1], email)) {
      if (updates.points !== undefined) {
        const oldPtLog = Number(data[i][10]) || 0;   // 変更前ポイント
        const rk     = getRank(updates.points);
        const nextMs = getNextMilestone(updates.points);
        const toNext = nextMs !== null ? (nextMs - updates.points) + "pt → " + getNextRankName(updates.points) : "🎖️ MAX (Fuji)";
        sh.getRange(i+1,11).setValue(updates.points);
        sh.getRange(i+1,12).setValue(rk.name);
        sh.getRange(i+1,13).setValue(rk.icon);
        sh.getRange(i+1,14).setValue(toNext);
        const rankColors = {Sakura:"#fce4ec",Ume:"#fbe9e7",Take:"#e8f5e9",Matsu:"#fff8e1",Fuji:"#e3f2fd"};
        sh.getRange(i+1, 1, 1, 17).setBackground(rankColors[rk.name] || "#ffffff");
        // ★ ポイント履歴を自動記録（全経路共通）
        if (updates.points !== oldPtLog) {
          logPointHistory(ss, data[i][0], data[i][1], data[i][2],
            oldPtLog, updates.points, rk.name, rk.icon,
            updates._source || "");
        }
      }
      if (updates.rank      !== undefined) sh.getRange(i+1,12).setValue(updates.rank);
      if (updates.lastVisit !== undefined) sh.getRange(i+1,15).setValue(updates.lastVisit);
      if (updates.history   !== undefined) {
        sh.getRange(i+1,16).setValue(updates.history.length);
        sh.getRange(i+1,17).setValue(updates.history.join(","));
      }
      return;
    }
  }
}

function ensureMember(ss, data) {
  const existing = findMember(data.email.toLowerCase());
  if (!existing) {
    // v7.0：handleRegister と同じシニア自動判定をここでも行う
    let memberType = data.memberType || "G";
    if (data.birthYear) {
      const age = new Date().getFullYear() - parseInt(data.birthYear);
      if (age >= 65) memberType = "SR";
    }
    saveMember(ss, {
      email:       data.email.toLowerCase(),
      name:        data.name,
      birthYear:   data.birthYear  || "",
      birthMonth:  data.birthMonth || "",
      birthDay:    data.birthDay   || "",
      language:    data.language   || "",
      gender:      data.gender     || "",
      memberType:  memberType,
      points:      0,
      rank:        "Sakura",
      registeredAt:Utilities.formatDate(new Date(),"Europe/Warsaw","yyyy-MM-dd"),
      lastVisit:   "",
      history:     []
    });
    // v6.8：注文アプリ内で新規登録した方にも登録完了メール（WhatsApp案内入り）を送る
    try {
      const saved = findMember(data.email.toLowerCase());
      if (saved) sendWelcomeMail(saved);
    } catch (e) { Logger.log("welcome mail (ensureMember): " + e); }
  }
}

// ════════════════════════════════════════════════════════════════════
//  注文シート書式
// ════════════════════════════════════════════════════════════════════
function createOrderSheet(ss) {
  const sh = ss.insertSheet(SHEET_ORDERS);
  const headers = [
    "Nr zamówienia / Order No.", "Data i godzina / Date & Time",
    "Stolik / Table", "Liczba gości / Guests", "Nr poz. / Item No.",
    "Nazwa PL / Name PL", "Nazwa EN / Name EN", "Ilość / Qty",
    "Cena jedn. / Unit price (zł)", "Suma poz. / Item total (zł)",
    "📋 MODYFIKACJE / MODIFICATIONS", "Uwagi klienta / Guest notes",
    "Członek / Member", "Imię i nazwisko / Full name",
    "Kontakt / Contact", "Miesiąc urodzin / Birth month",
    "Język / Language", "STATUS", "RAZEM / TOTAL (zł)",
    "商品ID / Item ID"
  ];
  sh.appendRow(headers);
  const hdr = sh.getRange(1,1,1,headers.length);
  hdr.setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e")
     .setHorizontalAlignment("center").setWrap(true).setFontSize(9);
  sh.getRange(1,11).setBackground("#7a1a0a").setFontColor("#ffffff").setFontWeight("bold");
  const widths = [165,155,90,90,65,185,185,55,110,115,280,230,110,165,185,120,85,125,115,110];
  widths.forEach((w,i) => sh.setColumnWidth(i+1, w));
  sh.setFrozenRows(1);
  return sh;
}

function formatOrderRows(sh, startRow, rowCount) {
  const count = Math.max(rowCount, 1);
  // v8.0：2行目に挿入するとヘッダー（濃色）の書式を引き継ぐため、必ずリセットする
  const w = Math.min(20, sh.getMaxColumns());
  sh.getRange(startRow, 1, count, w).clearFormat();
  // 注文の通し番号で縞模様を作る（新着が上に来ても交互になる）
  const color = (Math.floor((sh.getLastRow()-1)/count) % 2 === 0) ? "#fff8f0" : "#ffffff";
  sh.getRange(startRow,1,count,19).setBackground(color);
  sh.getRange(startRow,11,count,1).setWrap(true).setVerticalAlignment("top")
    .setBackground("#fffde7").setFontSize(9);
  sh.getRange(startRow,18,count,1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(["NOWE / NEW","W TRAKCIE / IN PROGRESS","GOTOWE / READY","WYDANE / SERVED"],true).build())
    .setFontColor("#c0392b").setFontWeight("bold");
  if (count > 0) sh.getRange(startRow,19).setFontWeight("bold").setFontColor("#c0392b");
  sh.autoResizeRows(startRow, count);
}

// ── 次のマイルストーン ────────────────────────────────────────────────
function getNextMilestone(pt) {
  const milestones = [50, 100, 200, 500];
  return milestones.find(m => pt < m) || null;
}

function getNextRankName(pt) {
  if (pt < 50)  return "Ume 🏅";
  if (pt < 100) return "Take 🥉";
  if (pt < 200) return "Matsu 🥈";
  if (pt < 500) return "Fuji 👑";
  return "MAX";
}

// ── ユーティリティ ──────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Orders sheet: " + (ss.getSheetByName(SHEET_ORDERS) ? "OK" : "Not found"));
  Logger.log("Members sheet: " + (ss.getSheetByName(SHEET_MEMBERS) ? "OK" : "Not found"));
}

// ════════════════════════════════════════════════════════════════════
//  ポイント失効ルール（8ヶ月ローリング）
//  GASトリガー：毎日深夜に checkExpiryAndNotify() を実行
// ════════════════════════════════════════════════════════════════════

const EXPIRY_MONTHS   = 8;   // 有効期限（月）
const REMIND_DAYS     = 14;  // リマインド（失効N日前）

// ── 毎日自動実行：失効チェック＋リマインドメール ─────────────────
function checkExpiryAndNotify() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sh      = getMemberSheet(ss);
  const data    = sh.getDataRange().getValues();
  const today   = new Date();

  for (let i = 1; i < data.length; i++) {
    const email     = (data[i][1] || "").toString().trim();
    const name      = (data[i][2] || "").toString();
    const points    = Number(data[i][10]) || 0;
    const rankName  = (data[i][11] || "Sakura").toString();
    const lastVisit = data[i][14];

    if (!email) continue;

    // 誕生日当日チェック（年・月・日すべてある場合）
    const birthYear  = (data[i][4] || "").toString();
    const birthMonth = (data[i][5] || "").toString();
    const birthDay   = (data[i][6] || "").toString();
    if (birthYear && birthMonth && birthDay) {
      const nowM = today.getMonth() + 1;
      const nowD = today.getDate();
      if (parseInt(birthMonth) === nowM && parseInt(birthDay) === nowD) {
        sendBirthdayMail(email, name, rankName);
        Logger.log("Birthday mail sent: " + email);
      }
    }

    if (!lastVisit || points === 0) continue;

    const lastDate   = new Date(lastVisit);
    const expiryDate = addMonths(lastDate, EXPIRY_MONTHS);
    const daysLeft   = Math.floor((expiryDate - today) / (1000*60*60*24));

    // 失効済み（0日以下）→ ポイントリセット
    if (daysLeft <= 0) {
      // 17列構成：K=累計pt(11) L=ランク(12) M=アイコン(13) N=次ランク(14)
      sh.getRange(i+1, 11).setValue(0);
      sh.getRange(i+1, 12).setValue("Sakura");
      sh.getRange(i+1, 13).setValue("🌸");
      sh.getRange(i+1, 14).setValue("50pt → Ume 🏅");
      sh.getRange(i+1, 1, 1, 17).setBackground("#fce4ec");
      // ポイント履歴に失効を記録
      logPointHistory(ss, data[i][0], email, name,
        points, 0, "Sakura", "🌸", "自動失効（8ヶ月） / Auto expiry (8mo)");
      // ※ 来店履歴（L列）は保持（リセットしない）
      // 失効通知メール
      sendExpiryMail(email, name, rankName, points);
      Logger.log("Expired: " + email);
      continue;
    }

    // 失効2週間前 → リマインドメール（ちょうど14日前の日のみ送信）
    if (daysLeft === REMIND_DAYS) {
      sendReminderMail(email, name, rankName, points, expiryDate);
      Logger.log("Reminder sent: " + email + " (days left: " + daysLeft + ")");
    }
  }
}

// ── 失効2週間前リマインドメール ──────────────────────────────────
function sendReminderMail(email, name, rankName, points, expiryDate) {
  const expStr = Utilities.formatDate(expiryDate, "Europe/Warsaw", "dd.MM.yyyy");
  const rank   = RANKS.find(r => r.name === rankName) || RANKS[0];

  const subject = "🍜 Twoje punkty Wabi Navi wkrótce wygasną! / Your Wabi Navi points expire soon!";

  const body = `
Drogi/a ${name}様,

${rank.icon} ${rankName}ランクの特典がまもなく失効します！
Twoje korzyści z rankingu ${rankName} wkrótce wygasną!
Your ${rankName} rank benefits will expire soon!

━━━━━━━━━━━━━━━━━━━━━━━━━━
現在のポイント / Punkty: ${points} pt
失効日 / Data wygaśnięcia: ${expStr}
━━━━━━━━━━━━━━━━━━━━━━━━━━

今週、温かいラーメンはいかがですか？
W tym tygodniu zapraszamy na gorące ramen!
How about a warm bowl of ramen this week?

ご来店でポイントが8ヶ月延長されます。
Odwiedź nas — punkty zachowają ważność przez kolejne 8 miesięcy!
Visit us to keep your points valid for another 8 months!

📍 Wabi Navi
   ul. Małe Garbary 5A
   Warszawa

🌐 ${SITE_URL}

━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI Members Club
このメールはWabi Navi会員向けの自動送信メールです。
Ten e-mail został wysłany automatycznie do członków klubu Wabi Navi.
━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  try {
    GmailApp.sendEmail(email, subject, body);
  } catch(e) {
    Logger.log("Mail error: " + e.toString());
  }
}

// ── ポイント失効通知メール ────────────────────────────────────────
function sendExpiryMail(email, name, rankName, points) {
  const subject = "🌸 Wabi Navi ポイントが失効しました / Punkty wygasły";

  const body = `
Drogi/a ${name}様,

8ヶ月間のご来店がなかったため、ポイントが失効しました。
Twoje punkty wygasły z powodu braku wizyty przez 8 miesięcy.
Your points have expired due to no visit for 8 months.

━━━━━━━━━━━━━━━━━━━━━━━━━━
失効ポイント / Wygasłe punkty: ${points} pt (${rankName})
現在のポイント / Obecne punkty: 0 pt (Sakura 🌸)
━━━━━━━━━━━━━━━━━━━━━━━━━━

またのご来店を心よりお待ちしております！
Zapraszamy ponownie — zawsze miło Cię gościć!
We look forward to welcoming you back!

新しくポイントを貯め始めることができます。
Możesz zacząć zbierać punkty od nowa.

📍 Wabi Navi — ul. Małe Garbary 5A, 87-100 Toruń

━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI Members Club
━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  try {
    GmailApp.sendEmail(email, subject, body);
  } catch(e) {
    Logger.log("Expiry mail error: " + e.toString());
  }
}

// ── 誕生日当日メール ─────────────────────────────────────────────
function sendBirthdayMail(email, name, rankName) {
  const rank    = RANKS.find(r => r.name === rankName) || RANKS[0];
  const subject = "🎂 お誕生日おめでとうございます！ / Wszystkiego najlepszego! / Happy Birthday!";
  const body = `
${name}様 / Drogi/a ${name},

🎂 お誕生日おめでとうございます！
Wszystkiego najlepszego z okazji urodzin!
Happy Birthday!

━━━━━━━━━━━━━━━━━━━━━━━━━━
🎁 誕生日特典 / Prezent urodzinowy / Birthday gift:
ドリンク1杯無料！（コーヒー/紅茶/レモネード）
Napój gratis! (kawa/herbata/lemoniada)
Free drink! (coffee/tea/lemonade)

現在のランク / Twój ranking: ${rank.icon} ${rankName}
━━━━━━━━━━━━━━━━━━━━━━━━━━

今日はWabi Naviで特別なひとときを！
Zapraszamy dziś do Wabi Navi na wyjątkową chwilę!
Come celebrate your birthday with us today!

📍 Wabi Navi — ul. Małe Garbary 5A, 87-100 Toruń

━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI Members Club
━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  try {
    GmailApp.sendEmail(email, subject, body);
  } catch(e) {
    Logger.log("Birthday mail error: " + e.toString());
  }
}

// ── 日付ユーティリティ ────────────────────────────────────────────
function addMonths(date, months) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

// ── スタンプ時の失効チェック（来店時にも確認） ───────────────────
function checkMemberExpiry(member) {
  if (!member.lastVisit || member.points === 0) return false;
  const lastDate   = new Date(member.lastVisit);
  const expiryDate = addMonths(lastDate, EXPIRY_MONTHS);
  const today      = new Date();
  return today > expiryDate; // trueなら失効
}

// ── GASトリガー設定（初回1回だけ実行） ───────────────────────────
function setupDailyTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkExpiryAndNotify') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎日深夜1時に実行
  ScriptApp.newTrigger('checkExpiryAndNotify')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  Logger.log("✅ Daily trigger set: checkExpiryAndNotify @ 01:00");
}

// ════════════════════════════════════════════════════════════════════
//  係員用：スプレッドシート上のワンクリックポイント付与
//  【設定手順】
//  1. Apps Script エディタで「setupStaffSheet()」を1回実行
//  2. Spreadsheetに「係員操作 / Staff」シートが自動作成される
//  3. そのシートのE列「＋10pt」ボタンをクリックするだけで付与
// ════════════════════════════════════════════════════════════════════

const SHEET_STAFF = "係員操作 / Staff";
const SHEET_POINTLOG = "ポイント履歴 / Point Log";

// ════════════════════════════════════════════════════════════════════
//  メニュー管理（案A）：スプレッドシートがメニューの正本
//  【初回セットアップ】setupMenuSheet() を1回実行
//  【日常運用】シートのセルを編集するだけでアプリに反映（最大1分）
//    G列「状態」: 表示 / 売り切れ / 非表示
//    F列「価格」: 数字を変更すれば価格改定
//    行の並び順 = アプリでの表示順（行の移動で順番変更可）
// ════════════════════════════════════════════════════════════════════

const SHEET_MENU = "メニュー / Menu";

// ════════════════════════════════════════════════════════════════════
//  書籍貸し出し（会員特典）
//  【初回】setupBooksSheet() を1回実行 → 「書籍 / Books」シート生成
//  【日常】蔵書追加=行を追加 / 返却=F列を「在庫あり」に戻すだけ
//          （貸出情報G〜J列は自動クリア）
// ════════════════════════════════════════════════════════════════════

const SHEET_BOOKS = "書籍 / Books";
const BOOK_MAX_DAYS = 21;   // 最大貸出期間（3週間）

function setupBooksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_BOOKS);
  if (sh) { SpreadsheetApp.getUi().alert("「書籍 / Books」シートは既に存在します。"); return; }
  sh = ss.insertSheet(SHEET_BOOKS);
  const headers = ["書籍ID","タイトル / Tytuł","著者 / Autor","言語","ジャンル",
                   "状態","貸出会員ID","会員名","貸出日","返却予定日","メモ"];
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#2c1810")
    .setFontColor("#c8a96e").setHorizontalAlignment("center").setFontSize(9);
  sh.setFrozenRows(1);
  // サンプル行（削除して実蔵書に差し替えてください）
  sh.appendRow(["b1","（例）ノルウェイの森 / Norwegian Wood","村上春樹 / Haruki Murakami","JP/PL","小説 / Powieść","在庫あり","","","","","サンプル行・削除可"]);
  sh.appendRow(["b2","（例）Sztuka herbaty — 茶の本","岡倉天心 / Kakuzō Okakura","PL","文化 / Kultura","在庫あり","","","","","サンプル行・削除可"]);
  // 状態プルダウン
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["在庫あり","貸出中","非表示"], true).setAllowInvalid(false).build();
  sh.getRange(2, 6, 500, 1).setDataValidation(rule).setHorizontalAlignment("center").setFontWeight("bold");
  // 条件付き書式：貸出中=橙 / 非表示=灰
  const range = sh.getRange(2, 1, 500, headers.length);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="貸出中"')
      .setBackground("#fdebd0").setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="非表示"')
      .setBackground("#e8e8e8").setFontColor("#999").setRanges([range]).build(),
  ]);
  [70,240,180,60,110,80,120,120,95,95,180].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  Logger.log("✅ 書籍シート作成完了");
}

// ── 書籍一覧取得（30秒キャッシュ） ────────────────────────────────
function getBooksList() {
  const cache = CacheService.getScriptCache();
  const c = cache.get("books_v1");
  if (c) return JSON.parse(c);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_BOOKS);
  if (!sh) return { status:"ok", books: [] };
  const data = sh.getDataRange().getValues();
  const books = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const id = (r[0]||"").toString().trim();
    if (!id) continue;
    const status = (r[5]||"在庫あり").toString();
    if (status.indexOf("非表示") !== -1) continue;
    const lent = status.indexOf("貸出中") !== -1;
    let dueDate = "";
    if (lent && r[9]) {
      dueDate = (r[9] instanceof Date)
        ? Utilities.formatDate(r[9], "Europe/Warsaw", "yyyy-MM-dd")
        : r[9].toString();
    }
    books.push({ id, title:(r[1]||"").toString(), author:(r[2]||"").toString(),
      lang:(r[3]||"").toString(), genre:(r[4]||"").toString(), lent, dueDate });
  }
  const result = { status:"ok", books };
  try { cache.put("books_v1", JSON.stringify(result), 30); } catch(e){}
  return result;
}

// ── 貸出処理 ──────────────────────────────────────────────────────
function handleBorrowBook(data) {
  const bookId     = (data.bookId||"").toString().trim();
  const email      = (data.email||"").toString().trim().toLowerCase();
  const returnDate = (data.returnDate||"").toString().trim();  // yyyy-MM-dd
  if (!bookId || !email || !returnDate)
    return jsonResponse({ status:"error", message:"入力が不足しています / Brak danych" });

  // 会員確認
  const member = findMember(email);
  if (!member)
    return jsonResponse({ status:"notMember",
      message:"会員登録が必要です / Wymagane członkostwo / Members only" });

  // 期間チェック（今日〜3週間以内）
  const today = new Date(); today.setHours(0,0,0,0);
  const due   = new Date(returnDate + "T00:00:00");
  const maxD  = new Date(today.getTime() + BOOK_MAX_DAYS*24*3600*1000);
  if (isNaN(due.getTime()) || due < today || due > maxD)
    return jsonResponse({ status:"error",
      message:"返却予定日は本日から3週間以内で選択してください / Max 3 tygodnie" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_BOOKS);
  if (!sh) return jsonResponse({ status:"error", message:"書籍シート未作成" });

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0]||"").toString().trim() === bookId) {
      const status = (rows[i][5]||"").toString();
      if (status.indexOf("貸出中") !== -1)
        return jsonResponse({ status:"alreadyLent",
          message:"この本は貸出中です / Ta książka jest wypożyczona",
          dueDate: rows[i][9] ? rows[i][9].toString() : "" });

      const todayStr = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
      sh.getRange(i+1, 6).setValue("貸出中");
      sh.getRange(i+1, 7).setValue(member.memberId || "");
      sh.getRange(i+1, 8).setValue(member.name || email);
      sh.getRange(i+1, 9).setValue(todayStr);        // 受付日=貸出日
      sh.getRange(i+1,10).setValue(returnDate);
      try { CacheService.getScriptCache().remove("books_v1"); } catch(e){}

      return jsonResponse({ status:"ok",
        title: (rows[i][1]||"").toString(),
        lendDate: todayStr, returnDate,
        memberName: member.name || "",
        message: "貸出登録が完了しました / Zarezerwowano" });
    }
  }
  return jsonResponse({ status:"error", message:"書籍が見つかりません / Nie znaleziono" });
}


const MENU_SEED = {"cats": [{"id": "onigiri", "namePL": "Onigiri i Dania z Ryżem", "nameEN": "Onigiri & Rice Bowls", "subcategorized": false}, {"id": "przystawki", "namePL": "Przystawki", "nameEN": "Appetizers", "subcategorized": false}, {"id": "ramen", "namePL": "Ramen", "nameEN": "Ramen", "subcategorized": true}, {"id": "sushi", "namePL": "Sushi 🚧", "nameEN": "Sushi 🚧 (Coming Soon)", "subcategorized": false}, {"id": "dodatki", "namePL": "Dodatki do Ramen", "nameEN": "Ramen Add-ons", "subcategorized": false}, {"id": "napoje", "namePL": "Napoje, Herbaty, Kawy i Desery", "nameEN": "Drinks, Teas, Coffees & Desserts", "subcategorized": false}, {"id": "alkohole", "namePL": "Alkohole", "nameEN": "Alcohol", "subcategorized": false}, {"id": "takeaway", "namePL": "Pojemniki na Wynos", "nameEN": "Takeaway Containers", "subcategorized": false}], "rows": [["onigiri", "o1", "Onigiri (1 szt.)", "Onigiri (1 pc.)", "おにぎり", 14, "表示", "Japońska przekąska ryżowa z nadzieniem lub posypką.", "Japanese rice ball with filling or topping.", "Nadzienie: a.wodorosty b.kimchi🌶️ c.mięso mielone d.surimi-mayo e.tuńczyk-mayo f.krewetki-mayo | Posypka: g.warzywa h.shiso i.bonito j.czosnek-chili🌶️ k.jajko-nori", "Filling: a.seaweed b.kimchi🌶️ c.minced pork d.surimi-mayo e.tuna-mayo f.shrimp-mayo | Topping: g.veggies h.shiso i.bonito j.garlic-chili🌶️ k.egg-nori", "1,3,4,6", 0, "", "[{\"name\": \"Nadzienie / Filling\", \"type\": \"change\", \"options\": [\"a. Wodorosty / Seaweed (standard)\", \"b. Kimchi 🌶️\", \"c. Mięso mielone / Minced pork\", \"d. Surimi-mayo\", \"e. Tuńczyk-mayo / Tuna-mayo\", \"f. Krewetki-mayo / Shrimp-mayo\"], \"priceDeltas\": [0, 0, 0, 0, 0, 0]}, {\"name\": \"Posypka / Topping\", \"type\": \"change\", \"options\": [\"g. Warzywa / Veggies (standard)\", \"h. Shiso czerwone / Red shiso\", \"i. Bonito/Katsuo\", \"j. Czosnek-chili 🌶️ / Garlic-chili\", \"k. Jajko-Nori / Egg-Nori\"], \"priceDeltas\": [0, 0, 0, 0, 0]}]", "", false], ["onigiri", "o1b", "Onigiri (2 szt.)", "Onigiri (2 pcs.)", "おにぎり(2個)", 26, "表示", "Japońska przekąska ryżowa — podwójna porcja, do wyboru 2 rodzaje.", "Japanese rice ball — double portion, choice of 2 kinds.", "Jak wyżej — wybierz 2 rodzaje", "As above — choice of 2 kinds", "1,3,4,6", 0, "", "", "", false], ["onigiri", "o2", "Wabi Inari Sushi (2 szt.)", "Wabi Inari Sushi (2 pcs.)", "いなり寿司", 28, "表示", "Wegańskie. Soczyste i słodkie sakiewki z tofu z ryżem sushi.", "Vegan. Juicy sweet fried tofu pockets with sushi rice.", "Tofu smażone, ryż sushi", "Fried tofu, sushi rice", "1,6", 0, "", "", "", false], ["onigiri", "o3", "Tamago Don", "Tamago Don", "たまご丼", 38, "表示", "Puszyste jajko, słodka cebula, szczypiorek i sezam na ryżu.", "Fluffy egg, sweet onion, chives and sesame over rice.", "Jajko, cebula, szczypiorek, sezam, ryż", "Egg, onion, chives, sesame, rice", "3,11", 0, "Sezam / Sesame:1 | Szczypiorek / Chives:1", "", "", false], ["onigiri", "o4", "Oyako Don", "Oyako Don", "親子丼", 42, "表示", "Kurczak w sosie sojowym z jajkiem na ryżu.", "Chicken in soy sauce simmered with egg over rice.", "Kurczak, jajko, cebula, sos sojowy, ryż", "Chicken, egg, onion, soy sauce, rice", "1,3,6", 0, "Cebula / Onion:1", "", "", false], ["onigiri", "o5", "Soboro Don", "Soboro Don", "そぼろ丼", 43, "表示", "Mielone mięso wieprzowe i jajko na ryżu (opcja z tofu).", "Minced pork and egg over rice (tofu option).", "Wieprzowina mielona, jajko, cebula, ryż", "Minced pork, egg, onion, rice", "3", 0, "Cebula / Onion:1", "[{\"name\": \"Białko / Protein\", \"type\": \"change\", \"options\": [\"Wieprzowina (standard)\", \"Tofu — bez zmian ceny / no change\"], \"priceDeltas\": [0, 0]}]", "", false], ["onigiri", "o6", "Chashu Don", "Chashu Don", "チャーシュー丼", 46, "表示", "Karkówka z cebulą i puszystym jajkiem na ryżu.", "Pork neck with onion and fluffy egg over rice.", "Karkówka, jajko, cebula, ryż", "Pork neck, egg, onion, rice", "3", 0, "Cebula / Onion:1 | Jajko / Egg:3", "", "", false], ["onigiri", "o7", "Potato Ebi Don", "Potato Ebi Don", "ポテト巻き海老丼", 49, "表示", "Chrupiące krewetki w ziemniakach z jajkiem na ryżu.", "Crispy potato-wrapped shrimp and egg over rice.", "Krewetki, ziemniaki, jajko, ryż", "Shrimp, potato, egg, rice", "1,2,3", 0, "Jajko / Egg:3", "", "", false], ["onigiri", "o8", "Una Don", "Una Don", "うな丼", 50, "表示", "Grillowany węgorz w słodkim sosie kabayaki na ryżu.", "Grilled eel in sweet kabayaki sauce over rice.", "Węgorz, sos kabayaki, ryż", "Eel, kabayaki sauce, rice", "4,6", 0, "", "", "", false], ["onigiri", "o9", "Yasai Kakiage Don", "Vegetable Tempura Rice Bowl", "野菜かき揚げ丼", 48, "表示", "Placuszki z warzyw w tempurze na ryżu.", "Crispy vegetable fritters over rice.", "Warzywa, tempura, ryż", "Vegetables, tempura batter, rice", "1", 0, "", "", "", false], ["onigiri", "o10", "Ryż Parowany", "Steamed Rice", "ライス", 12, "表示", "Mała miseczka ryżu parowanego.", "A small portion of fluffy steamed rice.", "Ryż japoński", "Japanese rice", "", 0, "", "", "", false], ["onigiri", "o11", "Zestaw Domowy", "Home Set", "ご飯とみそ汁セット", 18, "表示", "Ryż parowany i gorący bulion miso z wakame.", "Steamed rice and hot miso broth with wakame.", "Ryż, bulion miso, wakame", "Rice, miso broth, wakame", "6", 0, "", "", "", false], ["przystawki", "a1", "Gyoza — Warzywne", "Gyoza — Vegetable", "野菜餃子", 29, "表示", "Pierożki z warzywami (6 szt.).", "Vegetable dumplings (6 pcs).", "Warzywa, ciasto pszenne", "Vegetables, wheat dough", "1,6", 0, "", "", "", false], ["przystawki", "a1b", "Gyoza — Kurczak", "Gyoza — Chicken", "チキン餃子", 30, "表示", "Pierożki z kurczakiem (6 szt.).", "Chicken dumplings (6 pcs).", "Kurczak, ciasto pszenne", "Chicken, wheat dough", "1", 0, "", "", "", false], ["przystawki", "a1c", "Gyoza — Wołowina", "Gyoza — Beef", "牛肉餃子", 34, "表示", "Pierożki z wołowiną (6 szt.).", "Beef dumplings (6 pcs).", "Wołowina, ciasto pszenne", "Beef, wheat dough", "1", 0, "", "", "", false], ["przystawki", "a1d", "Gyoza — Dynia i Chili", "Gyoza — Pumpkin & Chili", "カボチャ餃子", 29, "表示", "Pierożki z dynią i chili — yarō pikantne (5 szt.).", "Pumpkin & chili dumplings — mildly spicy (5 pcs).", "Dynia, chili, ciasto pszenne", "Pumpkin, chili, wheat dough", "1", 1, "", "", "", false], ["przystawki", "a1e", "Gyoza — Kaczka", "Gyoza — Duck", "鴨肉餃子", 32, "表示", "Pierożki z kaczką (5 szt.).", "Duck dumplings (5 pcs).", "Kaczka, ciasto pszenne", "Duck, wheat dough", "1", 0, "", "", "", false], ["przystawki", "a1f", "Gyoza — Krewetki", "Gyoza — Shrimp", "エビ餃子", 32, "表示", "Pierożki z krewetkami (5 szt.).", "Shrimp dumplings (5 pcs).", "Krewetki, ciasto pszenne", "Shrimp, wheat dough", "1,2", 0, "", "", "", false], ["przystawki", "a1g", "Gyoza — Jabłko", "Gyoza — Apple", "リンゴ餃子", 27, "表示", "Słodkie pierożki z jabłkiem (5 szt.).", "Sweet apple dumplings (5 pcs).", "Jabłko, ciasto pszenne", "Apple, wheat dough", "1", 0, "", "", "", false], ["przystawki", "a2", "Sałatka Orientalna z Kalmarów", "Oriental Squid Salad", "イカ山菜サラダ", 34, "表示", "Kalmary z grzybami mung i bambusem.", "Squid with wood ear mushrooms and bamboo.", "Kalmary, grzyby mun, bambus", "Squid, wood ear mushrooms, bamboo", "14", 0, "", "", "", false], ["przystawki", "a3", "Edamame", "Edamame", "枝豆", 23, "表示", "Chrupiące nasiona soi (wegańskie).", "Crisp soybeans. Vegan.", "Soja, sól morska", "Soybeans, sea salt", "6", 0, "", "", "", false], ["przystawki", "a4", "Karaage", "Karaage", "唐揚げ", 32, "表示", "Chrupiące kawałki kurczaka polane sosem teriyaki.", "Crispy fried chicken with teriyaki sauce.", "Kurczak, sos teriyaki", "Chicken, teriyaki sauce", "1,6", 0, "", "", "", false], ["przystawki", "a5", "Takoyaki (5 szt.)", "Takoyaki (5 pcs)", "たこ焼き", 33, "表示", "Smażone kulki nadziewane ośmiornicą.", "Fried balls with octopus filling.", "Ośmiornica, ciasto, warzywa", "Octopus, batter, vegetables", "1,3,14", 0, "", "", "", false], ["przystawki", "a6", "Krewetki w Ziemniakach", "Shrimp & Potato Dumplings", "海老のポテト巻き", 38, "表示", "Soczyste krewetki owinięte ziemniakami.", "Juicy shrimp wrapped in potato strings.", "Krewetki, ziemniaki", "Shrimp, potato", "2", 0, "", "", "", false], ["przystawki", "a7", "Mabo Dofu", "Mabo Dofu", "麻婆豆腐", 34, "表示", "Smażone tofu z mieloną wieprzowiną — bardzo pikantne.", "Fried tofu with minced pork — very spicy.", "Tofu, wieprzowina mielona, chili", "Tofu, minced pork, chili", "6", 3, "", "", "", false], ["przystawki", "a8", "Sałatka z Hijiki i Edamame", "Hijiki Salad with Edamame", "ひじき&枝豆のサラダ", 25, "表示", "Japońska sałatka z glonów i soi.", "Japanese seaweed and soybean salad.", "Hijiki, edamame", "Hijiki seaweed, edamame", "6,14", 0, "", "", "", false], ["przystawki", "a9", "Sałatka z Tofu", "Tofu Salad", "豆腐サラダ", 32, "表示", "Kostki tofu ze świeżymi warzywami i sosem sezamowym.", "Tofu cubes with fresh vegetables and sesame dressing.", "Tofu, warzywa, sezam", "Tofu, vegetables, sesame", "6,11", 0, "", "", "", false], ["przystawki", "a10", "Smażone Krewetki Panko", "Fried Shrimp (Panko)", "エビフライ", 38, "表示", "Krewetki w japońskiej panierce (3 szt.).", "Crispy panko-breaded shrimp (3 pcs).", "Krewetki, panko", "Shrimp, panko breadcrumbs", "1,2", 0, "", "", "", false], ["przystawki", "a11", "Tempura Warzywna", "Vegetable Tempura", "野菜天ぷら", 40, "表示", "Selekcja sezonowych warzyw w tempurze.", "Selection of seasonal vegetables in tempura batter.", "Warzywa sezonowe, ciasto tempura", "Seasonal vegetables, tempura batter", "1", 0, "", "", "", false], ["przystawki", "a12", "Tempura z Krewetek i Warzyw", "Shrimp & Vegetable Tempura", "エビ天かき揚げ", 45, "表示", "Krewetki i warzywa w delikatnym cieście.", "Shrimp and vegetables in light batter.", "Krewetki, warzywa, ciasto tempura", "Shrimp, vegetables, tempura batter", "1,2", 0, "", "", "", false], ["przystawki", "a13", "Kimchi", "Kimchi", "キムチ", 23, "表示", "Tradycyjna pikantna kiszona kapusta koreańska — średnio pikantne.", "Traditional spicy fermented Korean cabbage — medium spicy.", "Kapusta pekińska, chili", "Napa cabbage, chili", "2", 2, "", "", "", false], ["przystawki", "a14", "Goma Wakame", "Goma Wakame", "ゴマわかめ", 23, "表示", "Orzeźwiająca sałatka z glonów wakame z sezamem.", "Refreshing wakame seaweed salad with sesame.", "Wakame, sezam", "Wakame, sesame", "11", 0, "", "", "", false], ["przystawki", "a15", "Natto", "Natto", "納豆", 21, "表示", "Tradycyjna fermentowana soja, zdrowy rarytas.", "Traditional fermented soybeans, a healthy delicacy.", "Soja fermentowana", "Fermented soybeans", "6", 0, "", "", "", false], ["przystawki", "a16", "Ryż Parowany (Przystawka)", "Steamed Rice (Side)", "ライス", 12, "表示", "Mała porcja ryżu japońskiego.", "Small portion of Japanese rice.", "Ryż", "Rice", "", 0, "", "", "", false], ["przystawki", "a17", "Zestaw z Ryżem i Zupą Miso", "Rice and Miso Soup Set", "ごはんと味噌汁セット", 18, "表示", "Ryż parowany oraz zupa miso.", "Steamed rice with miso soup.", "Ryż, zupa miso", "Rice, miso soup", "6", 0, "", "", "", false], ["przystawki", "a18", "Sos Sezamowy", "Sesame Sauce", "ごまだれ", 3, "表示", "Gęsty, aromatyczny sos sezamowy własnej roboty.", "House-made rich aromatic sesame sauce.", "Sezam, sos sojowy", "Sesame, soy sauce", "6,11", 0, "", "", "", false], ["ramen", "r1", "Assari", "Assari", "あっさり", 47, "表示", "Subtelny i delikatny smak.", "Subtle and delicate flavor.", "Baza: sos sojowy, hondashi | Dodatki: kurczak, pół jajka, surimi, imbir, por", "Base: soy sauce, hondashi | Toppings: chicken, half egg, surimi, ginger, leek", "1,3,4,6", 0, "Kurczak / Chicken:5 | Jajko / Egg:3 | Surimi:3 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Kurczak / Chicken\", \"type\": \"add\", \"options\": [\"Standard\", \"Dodatkowy kurczak / Extra chicken +5 zł\"], \"priceDeltas\": [0, 5]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "tori", false], ["ramen", "r2", "Tori Paitan", "Tori Paitan", "鶏パイタン", 49, "表示", "Esencjonalny bulion drobiowy, intensywny.", "Rich, intense chicken broth.", "Baza: sos sojowy, paitan | Dodatki: kurczak, grzyby shiitake, pak choi, menma, imbir, por", "Base: soy sauce, paitan | Toppings: chicken, shiitake, pak choi, menma, ginger, leek", "1,6", 0, "Kurczak / Chicken:5 | Shiitake:2 | Pak Choi:2 | Menma:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "tori", false], ["ramen", "r3", "Tori Kimchi", "Tori Kimchi", "鶏キムチ", 49, "表示", "Pikantny, rozgrzewający i bogaty w smaku.", "Spicy, warming and rich in flavor.", "Baza: sos sojowy, paitan, chili | Dodatki: kurczak, kimchi, naruto, wakame, imbir, por", "Base: soy sauce, paitan, chili | Toppings: chicken, kimchi, naruto, wakame, ginger, leek", "1,2,4,6", 1, "Kurczak / Chicken:5 | Kimchi:4 | Naruto:3 | Wakame:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "tori", false], ["ramen", "r4", "Wabi Navi", "Wabi Navi", "和美なび", 47, "表示", "Sygnowany naszą nazwą, najchętniej wybierany ramen. Delikatny i klarowny smak.", "Our signature ramen — delicate and clear, the most popular choice.", "Baza: sos sojowy | Dodatki: karkówka wieprzowa, pół jajka, kiełki, imbir i por", "Base: soy sauce | Toppings: pork neck, half egg, bean sprouts, ginger, leek", "1,3,6", 0, "Karkówka / Pork neck:5 | Jajko / Egg:3 | Kiełki / Sprouts:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Karkówka / Chashu\", \"type\": \"change\", \"options\": [\"Standard\", \"Podwójna / Double +5 zł\", \"Bez / None — bez zmian ceny\"], \"priceDeltas\": [0, 5, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r5", "Miso", "Miso", "味噌", 49, "表示", "Łagodny bulion o głębokim aromacie miso.", "Mild broth with deep miso aroma.", "Baza: pasta miso | Dodatki: karkówka, pół jajka, kiełki, kukurydza, masło, imbir, por", "Base: miso paste | Toppings: pork neck, half egg, sprouts, corn, butter, ginger, leek", "1,6,7", 0, "Karkówka / Pork neck:5 | Jajko / Egg:3 | Kiełki / Sprouts:2 | Kukurydza / Corn:1 | Masło / Butter:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r6", "Tantanmen", "Tantanmen", "担々麺", 49, "表示", "Aksamitny, orzechowy, pikantny.", "Velvety, nutty, spicy.", "Baza: orzechowa pasta chili | Dodatki: mięso mielone wieprzowe, szpinak, imbir, por", "Base: peanut chili paste | Toppings: minced pork, spinach, ginger, leek", "1,5,6", 2, "Mięso mielone / Minced pork:5 | Szpinak / Spinach:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r7", "Tantan Paitan", "Tantan Paitan", "濃厚担々麺", 52, "表示", "Aksamitny, esencjonalny bulion o orzechowym i pikantnym smaku.", "Velvety, rich broth with nutty spicy flavor.", "Baza: orzechowa pasta chili, paitan | Dodatki: mięso mielone, pak choi, menma, shiitake, sezam biały, imbir, por", "Base: peanut chili paste, paitan | Toppings: minced pork, pak choi, menma, shiitake, white sesame, ginger, leek", "1,5,6,11", 1, "Mięso mielone / Minced pork:5 | Pak Choi:2 | Menma:2 | Shiitake:2 | Sezam / Sesame:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r8", "Tantan Miso", "Tantan Miso", "担々味噌", 52, "表示", "Głęboki smak umami.", "Deep umami flavor.", "Baza: orzechowa pasta chili, miso | Dodatki: mięso mielone, pak choi, kukurydza, grzyby mun, imbir, por", "Base: peanut chili paste, miso | Toppings: minced pork, pak choi, corn, wood ear mushrooms, ginger, leek", "1,5,6", 1, "Mięso mielone / Minced pork:5 | Pak Choi:2 | Kukurydza / Corn:1 | Grzyby Mun:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r9", "Tantan Kokosowy", "Tantan Coconut", "担々麺ココナッツミルク風味", 52, "表示", "Bogate Tantanmen z mleczkiem kokosowym. Delikatna słodycz i pikantny finisz.", "Rich tantanmen with coconut milk. Gentle sweetness with a spicy finish.", "Dodatki: mięso mielone wieprzowe, pak choi, kukurydza, menma, naruto, imbir, por", "Toppings: minced pork, pak choi, corn, menma, naruto, ginger, leek", "1,4,5", 1, "Mięso mielone / Minced pork:5 | Pak Choi:2 | Kukurydza / Corn:1 | Menma:2 | Naruto:3 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r10", "Tonkotsu", "Tonkotsu", "豚骨", 52, "表示", "Esencjonalny, gęsty bulion wieprzowy gotowany przez wiele godzin.", "Rich, thick pork bone broth simmered for many hours.", "Baza: paitan, sos sojowy | Dodatki: karkówka, pół jajka, kiełki, pędy bambusa, grzyby mun, naruto, kukurydza, imbir, por", "Base: paitan, soy sauce | Toppings: pork neck, half egg, sprouts, bamboo shoots, wood ear mushrooms, naruto, corn, ginger, leek", "1,4,6", 0, "Karkówka / Pork neck:5 | Jajko / Egg:3 | Kiełki / Sprouts:2 | Pędy bambusa / Bamboo shoots:2 | Grzyby Mun:2 | Naruto:3 | Kukurydza / Corn:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Karkówka / Chashu\", \"type\": \"change\", \"options\": [\"Standard\", \"Podwójna / Double +5 zł\", \"Bez / None — bez zmian ceny\"], \"priceDeltas\": [0, 5, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r11", "Tonkotsu Miso", "Tonkotsu Miso", "豚骨味噌", 54, "表示", "Mocny wywar wieprzowy wzbogacony głębokim smakiem miso.", "Strong pork broth enriched with deep miso flavor.", "Baza: paitan, miso | Dodatki: karkówka, pół jajka, kiełki, szpinak, pędy bambusa, grzyby mun, naruto, imbir, por", "Base: paitan, miso | Toppings: pork neck, half egg, sprouts, spinach, bamboo shoots, wood ear mushrooms, naruto, ginger, leek", "1,4,6", 0, "Karkówka / Pork neck:5 | Jajko / Egg:3 | Kiełki / Sprouts:2 | Szpinak / Spinach:2 | Pędy bambusa / Bamboo shoots:2 | Grzyby Mun:2 | Naruto:3 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r12", "Mabo Ramen", "Mabo Ramen", "麻婆ラーメン", 54, "表示", "Wyrazisty duet pikanterii i umami.", "Bold duo of spice and umami.", "Baza: orzechowa pasta chili, mleko sojowe | Dodatki: Mabo Tofu (mięso mielone wieprzowe i tofu w aromatycznym sosie z nutą chili), shiitake, imbir, por", "Base: peanut chili paste, soy milk | Toppings: Mabo Tofu (minced pork and tofu in aromatic chili sauce), shiitake, ginger, leek", "1,5,6", 3, "Mabo Tofu:6 | Shiitake:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r13", "Wege Shoyu", "Vege Shoyu", "ベジ醤油", 46, "表示", "Delikatny, klarowny.", "Delicate and clear.", "Baza: sos sojowy | Dodatki: pół jajka, kiełki, szpinak, wakame, kukurydza, imbir, por", "Base: soy sauce | Toppings: half egg, sprouts, spinach, wakame, corn, ginger, leek", "1,3,6", 0, "Jajko / Egg:3 | Kiełki / Sprouts:2 | Szpinak / Spinach:2 | Wakame:2 | Kukurydza / Corn:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r14", "Wege Miso", "Vege Miso", "ベジ味噌", 48, "表示", "Wegański ramen z bazą miso.", "Vegan ramen with miso base.", "Baza: miso | Dodatki: pół jajka, kiełki, kukurydza, grzyby mun, imbir, por", "Base: miso | Toppings: half egg, sprouts, corn, wood ear mushrooms, ginger, leek", "1,3,6", 0, "Jajko / Egg:3 | Kiełki / Sprouts:2 | Kukurydza / Corn:1 | Grzyby Mun:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r15", "Wege Tantan", "Vege Tantan", "ベジ担々麺", 48, "表示", "Wegański tantanmen z kruszonym tofu.", "Vegan tantanmen with crumbled tofu.", "Baza: orzechowa pasta chili | Dodatki: kruszone tofu z warzywami, szpinak, kukurydza, pędy bambusa, imbir, por", "Base: peanut chili paste | Toppings: crumbled tofu with vegetables, spinach, corn, bamboo shoots, ginger, leek", "1,5,6", 2, "Tofu kruszone / Crumbled tofu:6 | Szpinak / Spinach:2 | Kukurydza / Corn:1 | Pędy bambusa / Bamboo shoots:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r16", "Wege Tantan Kokosowy", "Vege Tantan Coconut", "ベジ坦々麺ココナッツミルク味", 52, "表示", "Wegański, kokosowy tantanmen.", "Vegan coconut tantanmen.", "Baza: orzechowa pasta chili, mleko kokosowe | Dodatki: kruszone tofu z warzywami, pak choi, grzyby mun, kukurydza, imbir, por", "Base: peanut chili paste, coconut milk | Toppings: crumbled tofu with vegetables, pak choi, wood ear mushrooms, corn, ginger, leek", "1,5", 1, "Tofu kruszone / Crumbled tofu:6 | Pak Choi:2 | Grzyby Mun:2 | Kukurydza / Corn:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r17", "Wege Białe z Tofu", "Vege White with Tofu", "ベジ豆腐", 53, "表示", "Wegański bulion sojowy z kruszonym tofu.", "Vegan soy broth with crumbled tofu.", "Baza: sos sojowy, mleko sojowe | Dodatki: kruszone tofu, grzyby shiitake, pak choi, imbir, por", "Base: soy sauce, soy milk | Toppings: crumbled tofu, shiitake, pak choi, ginger, leek", "1,6", 0, "Tofu kruszone / Crumbled tofu:6 | Shiitake:2 | Pak Choi:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r18", "Wege Kimchi", "Vege Kimchi", "ベジキムチ", 48, "表示", "Pikantny bulion wegański z kimchi.", "Spicy vegan broth with kimchi.", "Baza: sos sojowy, chili | Dodatki: kimchi, pak choi, wakame, imbir, por", "Base: soy sauce, chili | Toppings: kimchi, pak choi, wakame, ginger, leek", "1,2,6", 1, "Kimchi:4 | Pak Choi:2 | Wakame:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "buta", false], ["ramen", "r19", "Seafood Ramen", "Seafood Ramen", "海鮮ラーメン", 52, "表示", "Delikatny, drobiowo-rybny.", "Delicate, poultry-fish broth.", "Baza: wywar z muli i przegrzebków, hondashi | Dodatki: krewetki, mule, wakame, naruto, nori, imbir, por", "Base: mussel & scallop stock, hondashi | Toppings: shrimp, mussels, wakame, naruto, nori, ginger, leek", "2,4,14", 0, "Krewetki / Shrimp:6 | Mule / Mussels:5 | Wakame:2 | Naruto:3 | Nori:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Extra owoce morza / Extra seafood\", \"type\": \"add\", \"options\": [\"Bez dodatku (standard)\", \"Extra porcja +50g (kalmary, ośmiornica, małże, krewetki) +6 zł\"], \"priceDeltas\": [0, 6]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r20", "Tantan z Krewetkami Kokosowy", "Tantan with Shrimp Coconut", "エビ入り担々麺ココナッツミルク味", 54, "表示", "Tantanmen kokosowy z krewetkami.", "Coconut tantanmen with shrimp.", "Baza: orzechowa pasta chili, mleko kokosowe | Dodatki: krewetki 3szt., pak choi, grzyby mun, naruto, goma wakame, kukurydza, imbir, por", "Base: peanut chili paste, coconut milk | Toppings: shrimp x3, pak choi, wood ear mushrooms, naruto, goma wakame, corn, ginger, leek", "1,2,4,5,11", 1, "Krewetki / Shrimp:6 | Pak Choi:2 | Grzyby Mun:2 | Naruto:3 | Goma Wakame:2 | Kukurydza / Corn:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r21", "Curry Tantanmen", "Curry Tantanmen", "カレー坦々麺", 53, "表示", "Połączenie curry i tantanmen.", "A fusion of curry and tantanmen.", "Baza: orzechowa pasta chili, curry, sos sojowy | Dodatki: mięso mielone wieprzowe, ser, szpinak, kukurydza, imbir, por", "Base: peanut chili paste, curry, soy sauce | Toppings: minced pork, cheese, spinach, corn, ginger, leek", "1,5,6,7", 3, "Mięso mielone / Minced pork:5 | Ser / Cheese:2 | Szpinak / Spinach:2 | Kukurydza / Corn:1 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r22", "Wege Curry Tantanmen", "Vege Curry Tantanmen", "ベジカレー坦々麺", 52, "表示", "Wegańska wersja curry tantanmen.", "Vegan version of curry tantanmen.", "Baza: orzechowa pasta chili, curry, sos sojowy | Dodatki: kruszone tofu z warzywami, szpinak, kukurydza, pędy bambusa, imbir, por", "Base: peanut chili paste, curry, soy sauce | Toppings: crumbled tofu with vegetables, spinach, corn, bamboo shoots, ginger, leek", "1,5,6", 3, "Tofu kruszone / Crumbled tofu:6 | Szpinak / Spinach:2 | Kukurydza / Corn:1 | Pędy bambusa / Bamboo shoots:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r23", "Wege Tempura Udon/Soba", "Vege Tempura Udon/Soba", "野菜天ぷらうどん/そば", 54, "表示", "Wybierz makaron UDON lub SOBA.", "Choose UDON or SOBA noodles.", "Baza: sos sojowy | Dodatki: sezonowe warzywa smażone w głębokim oleju, szpinak, imbir, por", "Base: soy sauce | Toppings: deep-fried seasonal vegetable tempura, spinach, ginger, leek", "1", 0, "Tempura warzywna / Vegetable tempura:4 | Szpinak / Spinach:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Makaron / Noodles\", \"type\": \"change\", \"options\": [\"Udon (standard)\", \"Soba — bez zmian ceny / no change\"], \"priceDeltas\": [0, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r24", "Udon/Soba z Grzybami", "Udon/Soba with Mushrooms", "きのこ山菜うどん/そば", 52, "表示", "Wybierz makaron UDON lub SOBA.", "Choose UDON or SOBA noodles.", "Baza: sos sojowy, konbudashi | Dodatki: nameko, smażone shiitake, pieczarki, grzyby słomkowe, naruto, szpinak, imbir, por", "Base: soy sauce, konbudashi | Toppings: nameko, fried shiitake, mushrooms, straw mushrooms, naruto, spinach, ginger, leek", "1,4", 0, "Nameko:2 | Shiitake:2 | Pieczarki / Mushrooms:2 | Grzyby słomkowe / Straw mushrooms:2 | Naruto:3 | Szpinak / Spinach:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Makaron / Noodles\", \"type\": \"change\", \"options\": [\"Udon (standard)\", \"Soba — bez zmian ceny / no change\"], \"priceDeltas\": [0, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r25", "Inari Udon/Soba", "Inari Udon/Soba", "いなりうどん/そば", 50, "表示", "Tradycyjny japoński bulion Kombu. Czysty, delikatny smak pełen naturalnego umami.", "Traditional Japanese kombu broth. Clean, delicate, full of natural umami.", "Dodatki: Inari — słodko-słone soczyste kieszonki z smażonego tofu, naruto, wakame, imbir, por", "Toppings: Inari — sweet-savory fried tofu pockets, naruto, wakame, ginger, leek", "1,4,6", 0, "Inari (kieszonki z tofu):3 | Naruto:3 | Wakame:2 | Imbir / Ginger:1 | Por / Leek:1", "[{\"name\": \"Makaron / Noodles\", \"type\": \"change\", \"options\": [\"Udon (standard)\", \"Soba — bez zmian ceny / no change\"], \"priceDeltas\": [0, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "oferta", false], ["ramen", "r26", "Salad Ramen", "Salad Ramen", "サラダラーメン", 49, "表示", "Makaron własnej roboty z sosem sezamowym.", "House-made noodles with sesame dressing.", "Sos sezamowy, pomidor, sałata lodowa, karkówka w kostkę, kukurydza, menma, pół jajka", "Sesame sauce, tomato, iceberg lettuce, diced pork neck, corn, menma, half egg", "1,3,11", 0, "Karkówka / Pork neck:5 | Jajko / Egg:3 | Kukurydza / Corn:1 | Menma:2 | Pomidor / Tomato:1 | Sałata / Lettuce:1", "[{\"name\": \"Białko / Protein\", \"type\": \"change\", \"options\": [\"Karkówka (standard)\", \"Kruszone tofu / Crumbled tofu (wegetariańska) — bez zmian ceny\"], \"priceDeltas\": [0, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "zimne", false], ["ramen", "r27", "Ten Zaru Udon/Soba", "Ten Zaru Udon/Soba", "冷やし野菜天ざるうどん又はそば", 50, "表示", "Zimny makaron z chrupiącą tempurą warzywną i sosem sezamowym.", "Cold noodles with crispy vegetable tempura and sesame sauce.", "Wybierz makaron UDON lub SOBA | Dodatki: nori, imbir", "Choose UDON or SOBA | Toppings: nori, ginger", "1,11", 0, "Tempura warzywna / Vegetable tempura:4 | Nori:1 | Imbir / Ginger:1", "[{\"name\": \"Makaron / Noodles\", \"type\": \"change\", \"options\": [\"Udon (standard)\", \"Soba — bez zmian ceny / no change\"], \"priceDeltas\": [0, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "zimne", false], ["ramen", "r28", "Zaru Soba lub Zaru Udon", "Zaru Soba or Zaru Udon", "ざるそば又はざるうどん", 40, "表示", "Lekki bulion solny z yuzu.", "Light salt broth with yuzu.", "Dodatki: chashu, pół jajka, naruto, por, skórka yuzu, zielenina mizuna", "Toppings: chashu, half egg, naruto, leek, yuzu zest, mizuna greens", "1,3,4", 0, "Chashu:5 | Jajko / Egg:3 | Naruto:3 | Por / Leek:1 | Yuzu:1 | Mizuna:1", "[{\"name\": \"Makaron / Noodles\", \"type\": \"change\", \"options\": [\"Soba (standard)\", \"Udon — bez zmian ceny / no change\"], \"priceDeltas\": [0, 0]}, {\"name\": \"🍜 Kaedama / 替え玉 (dodatkowa porcja makaronu)\", \"type\": \"add\", \"options\": [\"Bez / None (standard)\", \"Pół porcji / Half 半替え玉 +4 zł\", \"Pełna porcja / Full 替え玉 +7 zł\"], \"priceDeltas\": [0, 4, 7]}]", "zimne", false], ["sushi", "s1", "Sake Maki (6 szt.)", "Salmon Roll (6 pcs)", "サーモン巻き", 22, "表示", "Klasyczna rolka z łososiem. 🚧 Wkrótce dostępne!", "Classic salmon roll. 🚧 Coming soon!", "Ryż, łosoś, nori", "Rice, salmon, nori", "1,4", 0, "", "", "", true], ["sushi", "s2", "Maguro Maki (6 szt.)", "Tuna Roll (6 pcs)", "マグロ巻き", 22, "表示", "Rolka z tuńczykiem. 🚧 Wkrótce dostępne!", "Tuna roll. 🚧 Coming soon!", "Ryż, tuńczyk, nori", "Rice, tuna, nori", "1,4", 0, "", "", "", true], ["sushi", "s3", "Kappa Maki (6 szt.)", "Cucumber Roll (6 pcs)", "かっぱ巻き", 16, "表示", "Rolka z ogórkiem. Wegańska. 🚧 Wkrótce dostępne!", "Cucumber roll. Vegan. 🚧 Coming soon!", "Ryż, ogórek, nori", "Rice, cucumber, nori", "1", 0, "", "", "", true], ["sushi", "s4", "Futomaki (5 szt.)", "Futomaki (5 pcs)", "太巻き", 26, "表示", "Gruba rolka z wieloma składnikami. 🚧 Wkrótce dostępne!", "Thick roll with multiple fillings. 🚧 Coming soon!", "Ryż, warzywa, jajko, nori", "Rice, vegetables, egg, nori", "1,3", 0, "", "", "", true], ["sushi", "s5", "Uramaki California (8 szt.)", "California Uramaki (8 pcs)", "カリフォルニアロール", 28, "表示", "Rolka odwrócona z surimi i awokado. 🚧 Wkrótce dostępne!", "Inside-out roll with surimi and avocado. 🚧 Coming soon!", "Ryż, surimi, awokado, ogórek", "Rice, surimi, avocado, cucumber", "1,4", 0, "", "", "", true], ["sushi", "s6", "Uramaki Spicy Salmon (8 szt.)", "Spicy Salmon Uramaki (8 pcs)", "スパイシーサーモン", 30, "表示", "Rolka z pikantnym łososiem. 🚧 Wkrótce dostępne!", "Roll with spicy salmon. 🚧 Coming soon!", "Ryż, łosoś, majonez chili", "Rice, salmon, chili mayo", "1,4", 1, "", "", "", true], ["sushi", "s7", "Nigiri Sake (2 szt.)", "Salmon Nigiri (2 pcs)", "サーモン握り", 18, "表示", "Nigiri z łososiem. 🚧 Wkrótce dostępne!", "Salmon nigiri. 🚧 Coming soon!", "Ryż, łosoś", "Rice, salmon", "4", 0, "", "", "", true], ["sushi", "s8", "Nigiri Maguro (2 szt.)", "Tuna Nigiri (2 pcs)", "マグロ握り", 18, "表示", "Nigiri z tuńczykiem. 🚧 Wkrótce dostępne!", "Tuna nigiri. 🚧 Coming soon!", "Ryż, tuńczyk", "Rice, tuna", "4", 0, "", "", "", true], ["sushi", "s9", "Nigiri Ebi (2 szt.)", "Shrimp Nigiri (2 pcs)", "エビ握り", 18, "表示", "Nigiri z krewetką. 🚧 Wkrótce dostępne!", "Shrimp nigiri. 🚧 Coming soon!", "Ryż, krewetka", "Rice, shrimp", "2", 0, "", "", "", true], ["sushi", "s10", "Temaki Sake (1 szt.)", "Salmon Temaki (1 pc)", "サーモン手巻き", 16, "表示", "Ręcznie zwijana rolka z łososiem. 🚧 Wkrótce dostępne!", "Hand-rolled cone with salmon. 🚧 Coming soon!", "Ryż, łosoś, nori", "Rice, salmon, nori", "1,4", 0, "", "", "", true], ["dodatki", "d_noodle1", "Dodatkowa Porcja Makaronu", "Extra Noodles", "", 9, "表示", "Zwiększona porcja makaronu do ramenu.", "Larger portion of noodles for your ramen.", "Makaron pszenny", "Wheat noodles", "1", 0, "", "", "", false], ["dodatki", "d_noodle2", "Pół Porcji Makaronu (Kaedama)", "Half Noodles (Kaedama)", "", 6, "表示", "Dodatkowa pół-porcja makaronu — klasyczne kaedama!", "Extra half portion of noodles — classic kaedama refill!", "Makaron pszenny", "Wheat noodles", "1", 0, "", "", "", false], ["dodatki", "d_noodle3", "Zmiana na Makaron Ryżowy/Udon/Soba", "Change to Rice Noodle/Udon/Soba", "", 4, "表示", "Zamień standardowy makaron na inny rodzaj.", "Swap standard noodles for a different type.", "Makaron ryżowy / Udon / Soba", "Rice noodles / Udon / Soba", "1", 0, "", "", "", false], ["dodatki", "d_1a", "Szczypiorek", "Green Onions", "", 1, "表示", "Świeży szczypiorek.", "Fresh green onions.", "Szczypiorek", "Green onions", "", 0, "", "", "", false], ["dodatki", "d_1b", "Sezam", "Sesame Seeds", "", 1, "表示", "Prażony sezam.", "Toasted sesame seeds.", "Sezam", "Sesame", "11", 0, "", "", "", false], ["dodatki", "d_1c", "Nori", "Nori", "", 1, "表示", "Suszone wodorosty nori.", "Dried nori seaweed.", "Nori", "Nori", "", 0, "", "", "", false], ["dodatki", "d_1d", "Kukurydza", "Corn", "", 1, "表示", "Słodka kukurydza.", "Sweet corn.", "Kukurydza", "Corn", "", 0, "", "", "", false], ["dodatki", "d_1e", "Marynowany Imbir", "Pickled Ginger", "", 1, "表示", "Marynowany imbir.", "Pickled ginger.", "Imbir", "Ginger", "", 0, "", "", "", false], ["dodatki", "d_2a", "Pak Choi", "Pak Choi", "", 2, "表示", "Gotowany pak choi.", "Cooked pak choi.", "Pak choi", "Pak choi", "", 0, "", "", "", false], ["dodatki", "d_2b", "Szpinak", "Spinach", "", 2, "表示", "Świeży szpinak.", "Fresh spinach.", "Szpinak", "Spinach", "", 0, "", "", "", false], ["dodatki", "d_2c", "Wakame", "Wakame", "", 2, "表示", "Wodorosty wakame.", "Wakame seaweed.", "Wakame", "Wakame", "", 0, "", "", "", false], ["dodatki", "d_2d", "Czosnek", "Garlic", "", 2, "表示", "Świeży czosnek.", "Fresh garlic.", "Czosnek", "Garlic", "", 0, "", "", "", false], ["dodatki", "d_2e", "Grzyby Mun (Kikurage)", "Wood Ear Mushrooms", "", 2, "表示", "Czarne grzyby mun.", "Wood ear mushrooms.", "Grzyby mun", "Wood ear mushrooms", "", 0, "", "", "", false], ["dodatki", "d_2f", "Ser", "Cheese", "", 2, "表示", "Topiony ser.", "Melted cheese.", "Ser", "Cheese", "7", 0, "", "", "", false], ["dodatki", "d_3a", "Jajko (Połowa)", "Half-Egg", "", 3, "表示", "Połowa marynowanego jajka.", "Half marinated egg.", "Jajko", "Egg", "3", 0, "", "", "", false], ["dodatki", "d_3b", "Naruto", "Naruto", "", 3, "表示", "Tradycyjne ciasteczko rybne.", "Traditional fish cake.", "Surimi", "Surimi", "4", 0, "", "", "", false], ["dodatki", "d_4a", "Kimchi (Dodatek)", "Kimchi (Topping)", "", 4, "表示", "Pikantna kimchi jako dodatek.", "Spicy kimchi as a topping.", "Kapusta kimchi", "Kimchi cabbage", "2", 0, "", "", "", false], ["dodatki", "d_4b", "Hijiki z Edamame", "Hijiki with Edamame", "", 4, "表示", "Hijiki i edamame jako dodatek.", "Hijiki and edamame as a topping.", "Hijiki, edamame", "Hijiki, edamame", "6", 0, "", "", "", false], ["dodatki", "d_5a", "Wieprzowina Mielona", "Minced Pork", "", 5, "表示", "Mielona wieprzowina jako dodatek.", "Minced pork as a topping.", "Wieprzowina", "Pork", "", 0, "", "", "", false], ["dodatki", "d_5b", "Chashu (Karkówka)", "Chashu (Pork Neck)", "", 5, "表示", "Dodatkowy plaster chashu.", "Extra slice of chashu pork.", "Karkówka wieprzowa", "Pork neck", "", 0, "", "", "", false], ["dodatki", "d_5c", "Kurczak (Dodatek)", "Chicken (Topping)", "", 5, "表示", "Dodatkowy kurczak.", "Extra chicken.", "Kurczak", "Chicken", "", 0, "", "", "", false], ["dodatki", "d_6a", "Krewetki (3 szt.)", "Shrimp (3 pcs)", "", 6, "表示", "Dodatkowe krewetki.", "Extra shrimp.", "Krewetki", "Shrimp", "2", 0, "", "", "", false], ["dodatki", "d_6b", "Tofu Kruszone", "Minced Tofu", "", 6, "表示", "Kruszone tofu jako dodatek.", "Crumbled tofu as a topping.", "Tofu", "Tofu", "6", 0, "", "", "", false], ["dodatki", "d_6c", "Tofu Inari (2 szt.)", "Inari Tofu (2 pcs)", "", 6, "表示", "Słodkie sakiewki tofu.", "Sweet fried tofu pockets.", "Tofu inari", "Inari tofu", "6", 0, "", "", "", false], ["dodatki", "d_6d", "Dodatkowa Miseczka Ryżu", "Side Rice for Ramen", "", 6, "表示", "Mała miseczka ryżu do ramenu.", "Small bowl of rice for ramen.", "Ryż", "Rice", "", 0, "", "", "", false], ["dodatki", "d_6e", "Ekstra Owoce Morza", "Extra Seafood Topping", "", 6, "表示", "Dodatkowe owoce morza.", "Extra seafood mix.", "Owoce morza", "Seafood mix", "2,14", 0, "", "", "", false], ["napoje", "n1", "Choya Soda Yuzusshu 350ml", "Choya Soda Yuzusshu 350ml", "ゆずっしゅ", 19, "表示", "Orzeźwiający, gazowany napój z smakiem yuzu.", "Refreshing sparkling drink with yuzu flavour.", "Woda gazowana, yuzu", "Sparkling water, yuzu", "12", 0, "", "", "", false], ["napoje", "n2", "Choya Soda Umesshu 350ml", "Choya Soda Umesshu 350ml", "ウメッシュ", 19, "表示", "Orzeźwiający, gazowany napój z smakiem japońskiej śliwki.", "Refreshing sparkling drink with Japanese plum flavour.", "Woda gazowana, śliwka ume", "Sparkling water, ume plum", "12", 0, "", "", "", false], ["napoje", "n3", "Ryokucha 緑茶缶 340ml", "Ryokucha Green Tea Can 340ml", "緑茶缶", 17, "表示", "Zimna, orzeźwiająca zielona herbata w puszce bez cukru.", "Cold, refreshing green tea in a can, no added sugar.", "Zielona herbata", "Green tea", "", 0, "", "", "", false], ["napoje", "n4", "Matcha Latte 抹茶ラテ 280ml", "Matcha Latte 280ml", "抹茶ラテ", 20, "表示", "Bogaty smak matchy z kremowym mlekiem. Na ciepło lub zimno.", "Rich matcha flavour with creamy milk. Hot or cold.", "Matcha, mleko", "Matcha, milk", "7", 0, "", "[{\"name\": \"Temperatura\", \"type\": \"change\", \"options\": [\"Na ciepło / Hot (standard)\", \"Na zimno / Cold — bez zmian ceny\"], \"priceDeltas\": [0, 0]}]", "", false], ["napoje", "n5", "Pepsi ペプシ 200ml", "Pepsi 200ml", "ペプシ", 9, "表示", "Klasyczna cola Pepsi.", "Classic Pepsi cola.", "Napój gazowany", "Carbonated drink", "", 0, "", "", "", false], ["napoje", "n6", "Mirinda ミリンダ 200ml", "Mirinda 200ml", "ミリンダ", 9, "表示", "Oranżada Mirinda.", "Mirinda orange drink.", "Napój gazowany pomarańczowy", "Orange carbonated drink", "", 0, "", "", "", false], ["napoje", "n7", "Pepsi Max Zero ペプシ無糖 200ml", "Pepsi Max Zero 200ml", "ペプシ無糖", 9, "表示", "Pepsi bez cukru.", "Pepsi with no sugar.", "Napój gazowany bez cukru", "Sugar-free carbonated drink", "", 0, "", "", "", false], ["napoje", "n8", "Sok Jabłkowy りんごジュース 200ml", "Apple Juice 200ml", "りんごジュース", 9, "表示", "Sok jabłkowy.", "Apple juice.", "Sok jabłkowy", "Apple juice", "", 0, "", "", "", false], ["napoje", "n9", "Sok Pomarańczowy オレンジジュース 200ml", "Orange Juice 200ml", "オレンジジュース", 9, "表示", "Sok pomarańczowy.", "Orange juice.", "Sok pomarańczowy", "Orange juice", "", 0, "", "", "", false], ["napoje", "n10", "Woda Niegazowana ミネラルウォーター 300ml", "Still Mineral Water 300ml", "ミネラルウォーター(炭酸なし)", 9, "表示", "Woda mineralna niegazowana.", "Still mineral water.", "Woda mineralna", "Mineral water", "", 0, "", "", "", false], ["napoje", "n11", "Woda Gazowana ミネラルウォーター(炭酸入り) 300ml", "Sparkling Mineral Water 300ml", "ミネラルウォーター(炭酸入り)", 9, "表示", "Woda mineralna gazowana.", "Sparkling mineral water.", "Woda mineralna gazowana", "Sparkling mineral water", "", 0, "", "", "", false], ["napoje", "n12", "OKF Napój Gazowany Ananas スパークリング パイナップル 350ml", "OKF Sparkling Pineapple 350ml", "スパークリング パイナップル", 16, "表示", "Gazowany napój o smaku ananasowym.", "Sparkling drink with pineapple flavour.", "Woda gazowana, ananas", "Sparkling water, pineapple", "", 0, "", "", "", false], ["napoje", "n13", "OKF Napój Gazowany Brzoskwinia スパークリング ピーチ 350ml", "OKF Sparkling Peach 350ml", "スパークリング ピーチ", 16, "表示", "Gazowany napój o smaku brzoskwiniowym.", "Sparkling drink with peach flavour.", "Woda gazowana, brzoskwinia", "Sparkling water, peach", "", 0, "", "", "", false], ["napoje", "n14", "OKF Napój Gazowany Limonka スパークリング ライム 350ml", "OKF Sparkling Lime 350ml", "スパークリング ライム", 16, "表示", "Gazowany napój o smaku limonkowym.", "Sparkling drink with lime flavour.", "Woda gazowana, limonka", "Sparkling water, lime", "", 0, "", "", "", false], ["napoje", "n15", "OKF Vera King Aloes Granat ザクロ味 500ml", "OKF Vera King Aloe Pomegranate 500ml", "アロエ飲料 ザクロ味", 19, "表示", "Napój aloesowy z cząstkami aloesu, smak granatu.", "Aloe drink with aloe pieces, pomegranate flavour.", "Aloes, granat", "Aloe vera, pomegranate", "", 0, "", "", "", false], ["napoje", "n16", "OKF Vera King Aloes Kiwi キウィ味 500ml", "OKF Vera King Aloe Kiwi 500ml", "アロエ飲料 キウィ味", 19, "表示", "Napój aloesowy z cząstkami aloesu, smak kiwi.", "Aloe drink with aloe pieces, kiwi flavour.", "Aloes, kiwi", "Aloe vera, kiwi", "", 0, "", "", "", false], ["napoje", "n17", "Sencha 煎茶 24 zł", "Sencha Green Tea", "煎茶", 24, "表示", "Klasyczna japońska herbata o świeżym, roślinnym aromacie.", "Classic Japanese green tea with fresh, vegetal aroma.", "Herbata sencha", "Sencha tea", "", 0, "", "", "", false], ["napoje", "n18", "Genmaicha 玄米茶 24 zł", "Genmaicha", "玄米茶", 24, "表示", "Szlachetna zielona herbata z prażonym ryżem i matchą.", "Noble green tea with roasted rice and matcha.", "Herbata zielona, ryż prażony", "Green tea, roasted rice", "", 0, "", "", "", false], ["napoje", "n19", "Hojicha ほうじ茶 24 zł", "Hojicha", "ほうじ茶", 24, "表示", "Prażona japońska herbata o intensywnym, głębokim aromacie.", "Roasted Japanese tea with intense, deep aroma.", "Herbata hojicha", "Hojicha tea", "", 0, "", "", "", false], ["napoje", "n20", "Sencha Sakura 桜煎茶 24 zł", "Sencha Sakura", "桜煎茶", 24, "表示", "Aromatyczna herbata o zapachu japońskiej wiśni.", "Aromatic tea with Japanese cherry blossom scent.", "Sencha, płatki wiśni", "Sencha, cherry blossom", "", 0, "", "", "", false], ["napoje", "n21", "Sencha Cytrynowa レモン煎茶 24 zł", "Sencha Lemon", "レモン煎茶", 24, "表示", "Zielona herbata z cytryną i trawą cytrynową.", "Green tea with lemon and lemongrass.", "Sencha, cytryna", "Sencha, lemon", "", 0, "", "", "", false], ["napoje", "n22", "Sencha Jaśminowa ジャスミン茶 24 zł", "Sencha Jasmine", "ジャスミン茶", 24, "表示", "Zielona herbata z dodatkiem kwiatu jaśminu.", "Green tea with jasmine flower.", "Sencha, jaśmin", "Sencha, jasmine", "", 0, "", "", "", false], ["napoje", "n23", "Herbata Biała Pai Mu Tan Peach 白茶 24 zł", "White Tea Pai Mu Tan Peach", "白茶 白牡丹 パイムータンピーチ", 24, "表示", "Wyborna chińska herbata biała z brzoskwinią.", "Fine Chinese white tea with peach.", "Herbata biała, brzoskwinia", "White tea, peach", "", 0, "", "", "", false], ["napoje", "n24", "Herbata Oolong 烏龍茶 24 zł", "Oolong Tea", "烏龍茶", 24, "表示", "Chińska herbata oolong o głębokim aromacie.", "Chinese oolong tea with deep aroma.", "Herbata oolong", "Oolong tea", "", 0, "", "", "", false], ["napoje", "n25", "Pu-Erh Żurawina クランベリープーアル茶 24 zł", "Pu-erh Cranberry", "クランベリープーアル茶", 24, "表示", "Chińska herbata czerwona z żurawiną.", "Chinese red tea with cranberry.", "Pu-erh, żurawina", "Pu-erh, cranberry", "", 0, "", "", "", false], ["napoje", "n26", "Japońska Zima ジンジャーティー 24 zł", "Japanese Winter Ginger Tea", "ジンジャーティー", 24, "表示", "Zielona herbata z imbirem i miodem.", "Green tea with ginger and honey.", "Zielona herbata, imbir, miód", "Green tea, ginger, honey", "", 0, "", "", "", false], ["napoje", "n27", "Herbata Specjalna お任せホットティー 25 zł", "Special Tea (Chef's Choice)", "お任せホットティー", 25, "表示", "Herbata specjalna — zapytaj obsługę o smak.", "Special tea — ask staff for today's flavour.", "Herbata sezonowa", "Seasonal tea", "", 0, "", "", "", false], ["napoje", "n28", "Herbata na Zimno 14 zł", "Iced Tea", "アイスティー", 14, "表示", "Mango-Tango z cytrynowym orzeźwieniem. Na lodzie.", "Mango-Tango with citrus freshness. Served on ice.", "Herbata, mango, cytryna", "Tea, mango, lemon", "", 0, "", "", "", false], ["napoje", "n29", "Herbata Czarna w Torebce 12 zł", "Black Tea (Bag)", "紅茶ティーバッグ", 12, "表示", "Klasyczna czarna herbata w torebce.", "Classic black tea bag.", "Czarna herbata", "Black tea", "", 0, "", "", "", false], ["napoje", "n30", "Espresso エスプレッソ 12 zł", "Espresso", "エスプレッソ", 12, "表示", "Klasyczne espresso.", "Classic espresso.", "Kawa", "Coffee", "", 0, "", "", "", false], ["napoje", "n31", "Flat White フラットホワイト 14 zł", "Flat White", "フラットホワイト", 14, "表示", "Espresso z mlekiem.", "Espresso with milk.", "Kawa, mleko", "Coffee, milk", "7", 0, "", "", "", false], ["napoje", "n32", "Americano アメリカン 14 zł", "Americano", "アメリカン", 14, "表示", "Espresso z wodą.", "Espresso with water.", "Kawa, woda", "Coffee, water", "", 0, "", "", "", false], ["napoje", "n33", "Cappuccino カプチーノ 16 zł", "Cappuccino", "カプチーノ", 16, "表示", "Espresso z spienianym mlekiem.", "Espresso with foamed milk.", "Kawa, mleko", "Coffee, milk", "7", 0, "", "", "", false], ["napoje", "n34", "Latte Macchiato ラテマキアート 16 zł", "Latte Macchiato", "ラテマキアート", 16, "表示", "Mleko z espresso.", "Milk with espresso.", "Mleko, kawa", "Milk, coffee", "7", 0, "", "", "", false], ["napoje", "n35", "Mocha モカ 16 zł", "Mocha", "モカ", 16, "表示", "Espresso z czekoladą i mlekiem.", "Espresso with chocolate and milk.", "Kawa, czekolada, mleko", "Coffee, chocolate, milk", "7", 0, "", "", "", false], ["napoje", "n36", "Nesquik / KitKat Gorące Napoje 10 zł", "Kids Hot Drinks (Nesquik/KitKat)", "子供向けホットドリンク", 10, "表示", "Nesquik: kakao bez kofeiny. KitKat: czekoladowa chwila przyjemności.", "Nesquik: caffeine-free cocoa. KitKat: chocolate treat.", "Kakao / Czekolada, mleko", "Cocoa / Chocolate, milk", "7", 0, "", "", "", false], ["napoje", "n37", "Owocowe Gyoza z Lodami (3 szt.) 25 zł", "Fruit Gyoza with Ice Cream (3 pcs)", "りんご餃子3個＋お好みアイス1スクープ", 25, "表示", "Chrupiące pierożki z jabłkiem z galką ulubionych lodów.", "Crispy apple gyoza with a scoop of your favourite ice cream.", "Jabłko, ciasto, lody", "Apple, dough, ice cream", "1,7", 0, "", "[{\"name\": \"Lody / Ice cream\", \"type\": \"change\", \"options\": [\"Matcha 抹茶 (standard)\", \"Czarny Sezam 黒ゴマ — bez zmian ceny\", \"Wanilia バニラ — bez zmian ceny\"], \"priceDeltas\": [0, 0, 0]}]", "", false], ["napoje", "n38", "Lody Matcha 抹茶アイス 19 zł", "Matcha Ice Cream", "抹茶アイス", 19, "表示", "Lody z japońskiej zielonej herbaty matcha.", "Ice cream made from Japanese matcha green tea.", "Matcha, mleko, śmietana", "Matcha, milk, cream", "7", 0, "", "", "", false], ["napoje", "n39", "Lody Sezamowe 黒ゴマアイス 19 zł", "Black Sesame Ice Cream", "黒ゴマアイス", 19, "表示", "Lody o smaku czarnego sezamu.", "Ice cream with black sesame flavour.", "Czarny sezam, mleko, śmietana", "Black sesame, milk, cream", "7,11", 0, "", "", "", false], ["napoje", "n40", "Lody Waniliowe z Kokutō Umeshu バニラ黒糖梅酒がけ 21 zł", "Vanilla Ice Cream with Umeshu", "バニラ黒糖梅酒がけ", 21, "表示", "Lody waniliowe z likierem aromatycznej śliwki Ume.", "Vanilla ice cream with aromatic Ume plum liqueur.", "Wanilia, umeshu, mleko", "Vanilla, umeshu, milk", "7,12", 0, "", "", "", false], ["alkohole", "al1", "The Choya Ume Salute 200ml (Mała butelka)", "The Choya Ume Salute 200ml (Small bottle)", "チョーヤ スパークリング梅酒", 24, "表示", "Premium musujące wino z japońskich śliwek Ume. (Alk. 5,5%)", "Premium sparkling wine from Japanese Ume plums. (Alc. 5.5%)", "Śliwka ume, wino", "Ume plum, wine", "12", 0, "", "", "", false], ["alkohole", "al2", "Choya Sarari-to Shita Yuzu Shu — Kieliszek 100ml", "Choya Sarari Yuzu Shu — Glass 100ml", "チョーヤ さらりとした柚子酒 ロック", 24, "表示", "Japoński likier owocowy premium podawany na lodzie. (Alk. 7,5%)", "Premium Japanese fruit liqueur served on the rocks. (Alc. 7.5%)", "Yuzu, likier", "Yuzu, liqueur", "12", 0, "", "", "", false], ["alkohole", "al2b", "Choya Sarari-to Shita Yuzu Shu — Butelka 500ml", "Choya Sarari Yuzu Shu — Bottle 500ml", "チョーヤ さらりとした柚子酒 ボトル", 100, "表示", "Japoński likier yuzu — butelka 500ml. (Alk. 7,5%)", "Japanese yuzu liqueur — bottle 500ml. (Alc. 7.5%)", "Yuzu, likier", "Yuzu, liqueur", "12", 0, "", "", "", false], ["alkohole", "al3", "Choya Silver Red Wino — Kieliszek 125ml", "Choya Silver Red Wine — Glass 125ml", "チョーヤ シルバー 赤 フルーツワイン", 21, "表示", "Słodkie, czerwone wino śliwkowe. (Alk. 10%)", "Sweet red plum wine. (Alc. 10%)", "Śliwka, wino", "Plum, wine", "12", 0, "", "", "", false], ["alkohole", "al3b", "Choya Silver Red Wino — Butelka 500ml", "Choya Silver Red Wine — Bottle 500ml", "チョーヤ シルバー 赤 ボトル", 69, "表示", "Słodkie, czerwone wino śliwkowe — butelka. (Alk. 10%)", "Sweet red plum wine — bottle. (Alc. 10%)", "Śliwka, wino", "Plum, wine", "12", 0, "", "", "", false], ["alkohole", "al4", "Japońskie Sake Choya 日本酒 180ml", "Japanese Sake Choya 180ml", "チョーヤ 日本酒", 26, "表示", "Serwowane w tradycyjnej karafce Tokkuri. Lekko słodki. (Alk. 14,5%)", "Served in traditional Tokkuri carafe. Slightly sweet. (Alc. 14.5%)", "Sake", "Sake", "", 0, "", "[{\"name\": \"Temperatura\", \"type\": \"change\", \"options\": [\"Na ciepło / Warm (standard)\", \"Na zimno / Cold — bez zmian ceny\"], \"priceDeltas\": [0, 0]}]", "", false], ["alkohole", "al5", "Kirin Ichiban Lager z Beczki 生ビール 330ml", "Kirin Ichiban Lager Draft 330ml", "キリン一番搾り 生ビール", 22, "表示", "Najświeższe japońskie piwo prosto z beczki.", "Freshest Japanese beer straight from the tap.", "Piwo", "Beer", "1", 0, "", "", "", false], ["alkohole", "al5b", "Kirin Ichiban Lager z Beczki 生ビール 500ml", "Kirin Ichiban Lager Draft 500ml", "キリン一番搾り 生ビール 500ml", 25, "表示", "Duże piwo z beczki.", "Large draft beer.", "Piwo", "Beer", "1", 0, "", "", "", false], ["alkohole", "al6", "Asahi Super Dry アサヒ スーパードライ 330ml", "Asahi Super Dry 330ml", "アサヒ スーパードライ", 22, "表示", "Kultowa japońska Karakuchi. Czyste i orzeźwiające. (Alk. 5%)", "Iconic Japanese Karakuchi. Clean and refreshing. (Alc. 5%)", "Piwo", "Beer", "1", 0, "", "", "", false], ["alkohole", "al7", "Sapporo Premium Beer サッポロ プレミアム 330ml", "Sapporo Premium Beer 330ml", "サッポロ プレミアム", 22, "表示", "Eleganckie piwo premium o idealnie zbalansowanym smaku. (Alk. 4,8%)", "Elegant premium beer with perfectly balanced taste. (Alc. 4.8%)", "Piwo", "Beer", "1", 0, "", "", "", false], ["alkohole", "al8", "IKI Yuzu 330ml", "IKI Yuzu Beer 330ml", "IKI ゆず ビール", 22, "表示", "Orzeźwiające i cytrusowe. Zielona herbata i yuzu. (Alk. 4,5%)", "Refreshing and citrusy. Green tea and yuzu. (Alc. 4.5%)", "Piwo, yuzu, zielona herbata", "Beer, yuzu, green tea", "1", 0, "", "", "", false], ["alkohole", "al9", "IKI Ginger 330ml", "IKI Ginger Beer 330ml", "IKI ジンジャー ビール", 22, "表示", "Rozgrzewające i wyraziste. Z nutą imbiru. (Alk. 5,5%)", "Warming and distinctive. With a hint of ginger. (Alc. 5.5%)", "Piwo, imbir", "Beer, ginger", "1", 0, "", "", "", false], ["alkohole", "al10", "Tyskie トゥイスキエ 500ml", "Tyskie 500ml", "ポーランドビール Tyskie", 15, "表示", "Polskie piwo Tyskie.", "Polish Tyskie beer.", "Piwo", "Beer", "1", 0, "", "", "", false], ["alkohole", "al11", "Lech Pils レフ ピルス 500ml", "Lech Pils 500ml", "ポーランドビール Lech Pils", 15, "表示", "Polskie piwo Lech Pils.", "Polish Lech Pils beer.", "Piwo", "Beer", "1", 0, "", "", "", false], ["alkohole", "al12", "Kirin Ichiban Shibori キリン一番搾り ノンアルコール 330ml", "Kirin Ichiban Shibori Non-Alcoholic 330ml", "キリン一番搾り ノンアルコール", 19, "表示", "Pełnia smaku bez alkoholu. (Alk. 0,0%)", "Full flavour without alcohol. (Alc. 0.0%)", "Piwo bezalkoholowe", "Non-alcoholic beer", "1", 0, "", "", "", false], ["alkohole", "al13", "Lech Free レフ フリー ノンアルコール 330ml", "Lech Free Non-Alcoholic 330ml", "レフ フリー ノンアルコール", 14, "表示", "Bezalkoholowe Lech Free.", "Non-alcoholic Lech Free.", "Piwo bezalkoholowe", "Non-alcoholic beer", "1", 0, "", "", "", false], ["takeaway", "tk1", "Pojemnik Duży (Zupa/Ramen)", "Large Container (Soup/Ramen)", "", 2, "表示", "Pojemnik na wynos do zupy i ramenu.", "Takeaway container for soup and ramen.", "Pojemnik jednorazowy", "Disposable container", "", 0, "", "", "", false], ["takeaway", "tk2", "Pojemnik Mały (Przystawki)", "Small Container (Appetizers)", "", 1, "表示", "Pojemnik na wynos do przystawek.", "Takeaway container for appetizers.", "Pojemnik jednorazowy", "Disposable container", "", 0, "", "", "", false]]};

// ── 初回セットアップ：現行メニュー158品でシートを生成 ────────────
function setupMenuSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_MENU);
  if (sh) {
    const ui = SpreadsheetApp.getUi();
    const r = ui.alert("メニューシートは既に存在します",
      "作り直しますか？（現在の内容は消えます）", ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet(SHEET_MENU, 0);

  const headers = [
    "カテゴリーID","商品ID","商品名PL","商品名EN","商品名JP",
    "価格(zł)","状態","説明PL","説明EN","具材PL","具材EN",
    "アレルゲン","辛さ","抜ける具材（名前:減額単価 | 区切り）","オプションJSON(編集注意)",
    "だしベース","準備中"
  ];
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length).setFontWeight("bold")
    .setBackground("#2c1810").setFontColor("#c8a96e")
    .setHorizontalAlignment("center").setFontSize(9);
  sh.setFrozenRows(1); sh.setFrozenColumns(3);

  // データ書き込み
  const rows = MENU_SEED.rows;
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // G列：状態のプルダウン
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["表示","売り切れ","非表示"], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, 7, rows.length, 1).setDataValidation(rule)
    .setHorizontalAlignment("center").setFontWeight("bold");

  // 条件付き書式：売り切れ=橙 / 非表示=灰
  const range = sh.getRange(2, 1, rows.length, headers.length);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2="売り切れ"').setBackground("#fdebd0")
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2="非表示"').setBackground("#e8e8e8")
      .setFontColor("#999999").setRanges([range]).build(),
  ];
  sh.setConditionalFormatRules(rules);

  sh.setColumnWidth(1,90); sh.setColumnWidth(2,70); sh.setColumnWidth(3,220);
  sh.setColumnWidth(4,180); sh.setColumnWidth(5,120); sh.setColumnWidth(6,70);
  sh.setColumnWidth(7,90);  sh.setColumnWidth(8,260); sh.setColumnWidth(9,260);
  sh.setColumnWidth(10,180);sh.setColumnWidth(11,180);sh.setColumnWidth(12,90);
  sh.setColumnWidth(13,55); sh.setColumnWidth(14,220);sh.setColumnWidth(15,320);
  sh.setColumnWidth(16,90); sh.setColumnWidth(17,70);

  Logger.log("✅ メニューシート作成完了: " + rows.length + "品");
}

// ── シートからメニューJSONを構築（60秒キャッシュ） ────────────────
function getMenuFromSheet() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("menu_v1");
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) return null;   // 未セットアップ → アプリは内蔵メニューで動作

  const data = sh.getDataRange().getValues();
  const byCat = Object.create(null);

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const catId = (r[0]||"").toString().trim();
    const id    = (r[1]||"").toString().trim();
    if (!catId || !id) continue;
    const status = (r[6]||"表示").toString();
    if (status.indexOf("非表示") !== -1) continue;   // 非表示は配信しない

    const item = {
      id, imgKey:id,
      namePL:(r[2]||"").toString(), nameEN:(r[3]||"").toString(),
      nameJP:(r[4]||"").toString(), price:Number(r[5])||0,
      descPL:(r[7]||"").toString(), descEN:(r[8]||"").toString(),
      ingPL:(r[9]||"").toString(),  ingEN:(r[10]||"").toString(),
      allergens:(r[11]||"").toString().split(",").map(s=>s.trim()).filter(Boolean).map(Number),
      spice:Number(r[12])||0,
      removable: [],
      toppings:[], brothBase:(r[15]||"").toString() || undefined,
    };
    try { if (r[14]) item.toppings = JSON.parse(r[14]); } catch(e){}
    // N列「名前:減額単価」をパース（単価なし=減額0の飾り具材）
    (r[13]||"").toString().split("|").forEach(function(s){
      s = s.trim(); if (!s) return;
      const p  = s.split(":");
      const nm = p[0].trim();
      item.removable.push(nm);
      const v = Number(p[1]);
      if (v > 0) {
        if (!item.removableValues) item.removableValues = {};
        item.removableValues[nm] = v;
      }
    });
    if (r[16] === true || (r[16]||"").toString().toUpperCase()==="TRUE") item.comingSoon = true;
    if (status.indexOf("売り切れ") !== -1) item.soldOut = true;

    if (!byCat[catId]) byCat[catId] = [];
    byCat[catId].push(item);
  }

  const menu = MENU_SEED.cats.map(c => ({
    id:c.id, namePL:c.namePL, nameEN:c.nameEN,
    subcategorized:c.subcategorized, items: byCat[c.id] || []
  })).filter(c => c.items.length > 0);

  // ── ⭐おすすめカテゴリを先頭に差し込む（v6.0）──
  const promo = buildPromoCategory(ss, menu);
  if (promo) menu.unshift(promo);

  const result = { status:"ok", menu, updated: new Date().getTime() };
  try { cache.put("menu_v1", JSON.stringify(result), 60); } catch(e){}
  return result;
}





// ════════════════════════════════════════════════════════════════════
//  💬 会員用WhatsApp — 登録時の自動案内（v6.8）
//    ※WhatsAppのメッセージ自動送信にはBusiness API（審査・有料）が必要。
//      本実装は「参加リンクを登録直後に自動で届ける」方式：
//        ①登録完了メールにリンクを記載（メール登録の方）
//        ②登録完了画面に参加ボタンを表示（電話番号登録の方も含め全員）
// ════════════════════════════════════════════════════════════════════
function getWhatsAppUrl() {
  try { return PropertiesService.getScriptProperties().getProperty("WHATSAPP_URL") || ""; }
  catch (e) { return ""; }
}

function setupWhatsApp() {
  const ui  = SpreadsheetApp.getUi();
  const cur = getWhatsAppUrl();
  const r = ui.prompt("💬 WhatsApp dla członków / 会員用WhatsApp",
    (cur ? "Teraz / 現在:\n" + cur + "\n\n" : "Nie ustawiono / 未設定\n\n") +
    "Wklej link do kanału WhatsApp / 参加リンクを貼り付けてください:\n" +
    "  Kanał / チャンネル: https://whatsapp.com/channel/…\n" +
    "  Grupa / グループ:   https://chat.whatsapp.com/…\n" +
    "  Czat / 直接チャット: https://wa.me/48XXXXXXXXX\n\n" +
    "Aby usunąć wpisz / 削除するには:  OFF",
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const v = r.getResponseText().trim();
  const props = PropertiesService.getScriptProperties();
  if (v.toUpperCase() === "OFF" || v === "") {
    props.deleteProperty("WHATSAPP_URL");
    ui.alert("✅ Usunięto / 削除しました\nLink nie będzie już wysyłany.");
    return;
  }
  if (!/^https:\/\/(whatsapp\.com\/channel\/|chat\.whatsapp\.com\/|wa\.me\/)/.test(v)) {
    ui.alert("⚠️ Nieprawidłowy link / リンクの形式が正しくありません",
      "Link musi zaczynać się od:\n" +
      "  https://whatsapp.com/channel/…\n" +
      "  https://chat.whatsapp.com/…\n" +
      "  https://wa.me/…", ui.ButtonSet.OK);
    return;
  }
  props.setProperty("WHATSAPP_URL", v);
  ui.alert("✅ Zapisano / 保存しました",
    "Link będzie automatycznie wysyłany do nowych członków:\n" +
    "新規会員に自動でご案内されます。\n\n" +
    "・w e-mailu powitalnym / 登録完了メール\n" +
    "・przyciskiem po rejestracji / 登録完了画面のボタン\n\n" +
    v, ui.ButtonSet.OK);
}

// ── 登録完了メール（会員ID・特典・WhatsApp案内）────────────────────
function sendWelcomeMail(member) {
  const email = (member.email || "").trim();
  if (email.indexOf("@") === -1) return;      // 電話番号での登録はメール送信不可
  const wa   = getWhatsAppUrl();
  const name = member.name || "";
  const mid  = member.memberId || "";

  const subject = "Witamy w Wabi Navi Members Club! / Welcome! / ご入会ありがとうございます";
  const body = `
${name ? "Drogi/a " + name + "," : "Drogi/a Członku,"}

━━━━━━━━━━━━━━━━━━━━━━━━━━
🌸 Witamy w Wabi Navi Members Club!
   Welcome to the Wabi Navi Members Club!
   Wabi Navi Members Clubへようこそ。
━━━━━━━━━━━━━━━━━━━━━━━━━━

🪪 Twój numer członkowski / Your member ID / 会員番号
   ${mid}

Pokaż ten numer (lub kod QR z karty) przy kasie, aby zebrać punkty.
Show this ID (or the QR code on your card) at the counter to collect points.
レジでこの番号または会員カードのQRをご提示ください。

━━━━━━━━━━━━━━━━━━━━━━━━━━
🍜 Jak zbierać punkty / How it works / ポイントの貯め方
━━━━━━━━━━━━━━━━━━━━━━━━━━
• 1 miska ramenu = 10 pkt（1 pieczątka na wizytę）
  1 bowl of ramen = 10 points (1 stamp per visit)
  ラーメン1杯＝10ポイント（1来店1スタンプ）

• 50 pkt → ramen za pół ceny / half-price ramen / ラーメン半額
• 100 pkt → ramen gratis / free ramen / ラーメン1杯無料
  …i dalej co 50 pkt, na przemian / 以降50ptごとに交互

🎁 Bonus rangi: im wyższa ranga, tym więcej darmowych dodatków.
   Rank bonus: higher ranks get more free toppings.
   ランクが上がるほど無料トッピングが増えます。

🎂 W miesiącu urodzin: napój gratis!
   A free drink during your birthday month!
   誕生月はドリンク1杯無料！

━━━━━━━━━━━━━━━━━━━━━━━━━━
💳 Twoja karta członkowska / Your member card / 会員カード
━━━━━━━━━━━━━━━━━━━━━━━━━━
${APP_BASE}card.html
${wa ? `
━━━━━━━━━━━━━━━━━━━━━━━━━━
💬 Dołącz do nas na WhatsApp / Join us on WhatsApp
━━━━━━━━━━━━━━━━━━━━━━━━━━
Nowości, oferty specjalne i wydarzenia — prosto na Twój telefon.
News, special offers and events — straight to your phone.
最新情報・特別オファー・イベント情報をお届けします。

👉 ${wa}
` : ""}
━━━━━━━━━━━━━━━━━━━━━━━━━━
Dziękujemy i do zobaczenia wkrótce!
Thank you — see you soon!
またのご来店をお待ちしております。

📍 Wabi Navi, Toruń
━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI
━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  try { GmailApp.sendEmail(email, subject, body); }
  catch (e) { Logger.log("Welcome mail error: " + e.toString()); }
}

// ════════════════════════════════════════════════════════════════════
//  🗾 日本なんでもクイズ（待ち時間対策）— v6.7
//    設計方針：
//      ・1問正解 = 1pt（ラーメン1杯10ptに対し控えめ＝ポイント経済を守る）
//      ・1日3問まで（既定・変更可）＝ 1日最大3pt
//      ・その日に注文がある会員のみ（来店していない人は挑戦できない）
//      ・一度出た問題は同じ会員には二度出ない
//    初期値は OFF。メニューからいつでも切替。
// ════════════════════════════════════════════════════════════════════
const SHEET_TRIVIA    = "日本クイズ / Quiz o Japonii";
const SHEET_TRIVIALOG = "クイズ記録 / Quiz log";

function tvGet(k, def) {
  try { return PropertiesService.getScriptProperties().getProperty(k) || def; }
  catch (e) { return def; }
}
function triviaOn()    { return tvGet("TRIVIA_MODE", "off") === "on"; }
function triviaDaily() { return Number(tvGet("TRIVIA_DAILY", "3")) || 3; }
function triviaPoint() { return Number(tvGet("TRIVIA_POINT", "1")) || 1; }

// ── 設定（メニューから） ──────────────────────────────────────────
function setupTriviaSettings() {
  const ui = SpreadsheetApp.getUi();
  const cur = "Teraz / 現在: " + (triviaOn() ? "ON" : "OFF") +
              " | " + triviaDaily() + " pytań dziennie / 1日" + triviaDaily() + "問" +
              " | +" + triviaPoint() + " pkt za poprawną / 正解1問" + triviaPoint() + "pt";

  const r = ui.prompt("🗾 Quiz o Japonii / 日本なんでもクイズ",
    cur + "\n\n" +
    "Wpisz / 入力:  ON/OFF, pytań dziennie, punkty za poprawną\n" +
    "（ON/OFF, 1日の出題数, 正解1問あたりのポイント）\n\n" +
    "Przykład / 例:  ON,3,1\n" +
    "  → 3 pytania dziennie, 1 pkt za każdą poprawną odpowiedź\n" +
    "  （1日3問・正解1問につき1pt = 1日最大3pt）\n\n" +
    "Aby wyłączyć / 止めるには:  OFF",
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const p = r.getResponseText().split(",").map(x => x.trim());
  const props = PropertiesService.getScriptProperties();
  if ((p[0] || "").toUpperCase() === "OFF") {
    props.setProperty("TRIVIA_MODE", "off");
    ui.alert("✅ Quiz: OFF\nクイズを停止しました。");
    return;
  }
  if ((p[0] || "").toUpperCase() !== "ON") {
    ui.alert("⚠️ Zacznij od ON albo OFF / 最初に ON か OFF を入力してください");
    return;
  }
  const daily = Number(p[1]) || 3;
  const point = Number(p[2]) || 1;
  if (daily < 1 || daily > 10) { ui.alert("⚠️ Pytań dziennie: 1–10 / 出題数は1〜10"); return; }
  if (point < 1 || point > 5)  { ui.alert("⚠️ Punkty: 1–5 / ポイントは1〜5"); return; }

  props.setProperty("TRIVIA_MODE", "on");
  props.setProperty("TRIVIA_DAILY", String(daily));
  props.setProperty("TRIVIA_POINT", String(point));

  ui.alert("✅ Zapisano / 保存しました",
    "Quiz: ON\n" + daily + " pytań dziennie / 1日" + daily + "問\n" +
    "+" + point + " pkt za poprawną / 正解1問" + point + "pt\n" +
    "Maks. " + (daily * point) + " pkt dziennie / 1日最大" + (daily * point) + "pt\n\n" +
    "Tylko dla gości, którzy dziś złożyli zamówienie.\n当日注文のある会員のみ挑戦できます。",
    ui.ButtonSet.OK);
}

// ── 問題シートの作成（サンプル10問つき） ────────────────────────────
function setupTriviaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let sh = ss.getSheetByName(SHEET_TRIVIA);
  if (sh) {
    const r = ui.alert("Arkusz quizu już istnieje / クイズシートが既にあります",
      "Utworzyć od nowa? / 作り直しますか？（現在の問題は消えます）", ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet(SHEET_TRIVIA);
  const headers = ["ID", "Aktywne / 有効", "Pytanie PL", "Pytanie EN",
                   "Odp. 1", "Odp. 2", "Odp. 3", "Odp. 4",
                   "Poprawna 1-4 / 正解", "Wyjaśnienie PL", "Wyjaśnienie EN"];
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1f6f8b")
    .setFontColor("#ffffff").setHorizontalAlignment("center").setFontSize(9);
  [60, 90, 320, 320, 150, 150, 150, 150, 90, 300, 300]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);

  const q = [
    ["q1", true, "Jak po japońsku mówi się przed jedzeniem?", "What do Japanese people say before eating?",
     "Arigatō", "Itadakimasu", "Konnichiwa", "Sayōnara", 2,
     "„Itadakimasu” to podziękowanie za posiłek i wszystkim, którzy go przygotowali.",
     "“Itadakimasu” expresses gratitude for the meal and everyone who made it possible."],
    ["q2", true, "Co oznacza słowo „ramen”?", "What does the word “ramen” refer to?",
     "Rodzaj ryżu", "Zupa z makaronem", "Sos sojowy", "Herbata", 2,
     "Ramen to danie z pszennego makaronu w bulionie — nie rodzaj ryżu.",
     "Ramen is wheat noodles served in broth — not a type of rice."],
    ["q3", true, "Czy głośne siorbanie makaronu jest w Japonii niegrzeczne?", "Is slurping noodles rude in Japan?",
     "Tak, bardzo", "Nie — to naturalne", "Tylko w restauracjach", "Tylko w domu", 2,
     "Siorbanie schładza makaron i wzmacnia aromat. W Japonii jest całkowicie naturalne.",
     "Slurping cools the noodles and enhances aroma — it is perfectly normal in Japan."],
    ["q4", true, "Co to jest „kaedama”?", "What is “kaedama”?",
     "Dodatkowa porcja makaronu", "Rodzaj mięsa", "Miska", "Napój", 1,
     "Kaedama to dokładka samego makaronu do pozostałego bulionu.",
     "Kaedama is a refill of noodles added to your remaining broth."],
    ["q5", true, "Z czego robi się miso?", "What is miso made from?",
     "Z ryżu i wody", "Ze sfermentowanej soi", "Z ryb", "Z warzyw morskich", 2,
     "Miso to pasta ze sfermentowanej soi — podstawa japońskiej kuchni.",
     "Miso is a paste of fermented soybeans, a cornerstone of Japanese cooking."],
    ["q6", true, "Ile wysp tworzy główną część Japonii?", "How many main islands form Japan?",
     "2", "4", "7", "12", 2,
     "Cztery główne wyspy: Hokkaido, Honsiu, Sikoku i Kiusiu.",
     "Four main islands: Hokkaido, Honshu, Shikoku and Kyushu."],
    ["q7", true, "Co oznacza „umami”?", "What does “umami” mean?",
     "Słodycz", "Piąty smak — pełnia smaku", "Ostrość", "Kwaśność", 2,
     "Umami to piąty smak podstawowy, odkryty w Japonii w 1908 roku.",
     "Umami is the fifth basic taste, identified in Japan in 1908."],
    ["q8", true, "Czym jest „chashu” w ramenie?", "What is “chashu” in ramen?",
     "Duszona wieprzowina", "Jajko", "Wodorosty", "Kiełki", 1,
     "Chashu to długo duszona wieprzowina — klasyczny dodatek do ramenu.",
     "Chashu is slow-braised pork, a classic ramen topping."],
    ["q9", true, "Jak nazywa się japońskie kwiaty wiśni?", "What are Japanese cherry blossoms called?",
     "Momiji", "Sakura", "Matsu", "Take", 2,
     "Sakura — symbol wiosny i przemijania. To także pierwsza ranga w naszym klubie!",
     "Sakura — a symbol of spring. It is also the first rank in our members club!"],
    ["q10", true, "Co to jest „onigiri”?", "What is “onigiri”?",
     "Kulka ryżowa", "Zupa", "Makaron", "Deser", 1,
     "Onigiri to formowany ryż z nadzieniem — japońska przekąska numer jeden.",
     "Onigiri is a shaped rice ball with filling — Japan's number-one snack."],
  ];
  sh.getRange(2, 1, q.length, headers.length).setValues(q);
  sh.getRange(2, 2, q.length, 1).insertCheckboxes();

  // 記録シート
  let lg = ss.getSheetByName(SHEET_TRIVIALOG);
  if (!lg) {
    lg = ss.insertSheet(SHEET_TRIVIALOG);
    lg.appendRow(["Data / 日時", "ID członka", "Imię / Name", "Pytanie / 問題",
                  "Odpowiedź", "Wynik / Result", "Punkty"]);
    lg.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#1f6f8b")
      .setFontColor("#ffffff").setHorizontalAlignment("center").setFontSize(9);
    [150, 130, 150, 320, 80, 100, 70].forEach((w, i) => lg.setColumnWidth(i + 1, w));
    lg.setFrozenRows(1);
  }
  ui.alert("✅ Gotowe / 作成しました",
    "10 przykładowych pytań gotowych do użycia.\n" +
    "サンプル10問が入っています。自由に編集・追加してください。\n\n" +
    "・B列のチェックを外すと、その問題は出題されません\n" +
    "・I列は正解の番号（1〜4）\n\n" +
    "Włącz quiz w menu: 🗾 Quiz — ustawienia\nメニューの「🗾 Quiz」で公開できます。",
    ui.ButtonSet.OK);
}

// ── 当日の注文があるか（来店した人だけがクイズに挑戦できる）──────
function hasOrderedToday(ss, contact) {
  const sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) return false;
  const last = sh.getLastRow();
  if (last < 2) return false;
  // v8.0：注文は新しい順（2行目が最新）で記録されるため、上から800行を見る。
  //       旧版は末尾800行を読んでいたので、順序反転後は最も古い注文しか
  //       見えず、当日注文の判定（クイズの出題可否）が常に false になっていた。
  const rows = sh.getRange(2, 1, Math.min(last - 1, 800), 19).getValues();
  const todayStr = new Date().toDateString();
  const key = (contact || "").toString().trim().toLowerCase();
  const keyDigits = key.replace(/\D/g, "").slice(-9);

  for (let i = 0; i < rows.length; i++) {
    const t = rows[i][1];
    if (!(t instanceof Date) || t.toDateString() !== todayStr) continue;
    const c = (rows[i][14] || "").toString().trim().toLowerCase();
    if (!c) continue;
    if (c === key) return true;
    const cd = c.replace(/\D/g, "").slice(-9);
    if (keyDigits.length === 9 && cd === keyDigits) return true;
  }
  return false;
}

// ── GET action=getTrivia&memberId=… ──────────────────────────────
//    出題可能な問題（正解番号は含めない）と、本日の残り回数を返す
function getTriviaForMember(memberId) {
  if (!triviaOn()) return { status: "off" };
  const id = (memberId || "").toString().trim();
  if (!id) return { status: "error", message: "memberId required" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const member = findMember(id);
  if (!member) return { status: "notMember" };

  if (!hasOrderedToday(ss, member.email)) {
    return { status: "noOrder",
      message: "Quiz dostępny po złożeniu zamówienia / 本日ご注文のお客様が対象です" };
  }

  const sh = ss.getSheetByName(SHEET_TRIVIA);
  if (!sh) return { status: "noSheet" };

  // 本日の回答数と、これまでに answered した問題ID
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const lg = ss.getSheetByName(SHEET_TRIVIALOG);
  const myId = (member.memberId || "").toString().trim().toLowerCase();
  const doneIds = {};
  let todayCount = 0;
  if (lg) {
    const lr = lg.getDataRange().getValues();
    for (let i = 1; i < lr.length; i++) {
      if ((lr[i][1] || "").toString().trim().toLowerCase() !== myId) continue;
      const qid = (lr[i][3] || "").toString().split(" ")[0];
      if (qid) doneIds[qid] = true;
      const d = (lr[i][0] || "").toString().slice(0, 10);
      if (d === today) todayCount++;
    }
  }
  const remaining = Math.max(0, triviaDaily() - todayCount);
  if (remaining === 0) {
    return { status: "doneToday", remaining: 0, daily: triviaDaily(),
      message: "Dziś już wszystkie pytania! Zapraszamy jutro / 本日の分は終了しました" };
  }

  // 未回答の有効な問題から出題（毎回シャッフル）
  const rows = sh.getDataRange().getValues();
  const pool = [];
  for (let i = 1; i < rows.length; i++) {
    const qid = (rows[i][0] || "").toString().trim();
    const on  = rows[i][1] === true || (rows[i][1] || "").toString().toUpperCase() === "TRUE";
    if (!qid || !on || doneIds[qid]) continue;
    const opts = [rows[i][4], rows[i][5], rows[i][6], rows[i][7]]
      .map(x => (x || "").toString().trim()).filter(Boolean);
    if (opts.length < 2) continue;
    pool.push({ id: qid, pl: (rows[i][2] || "").toString(),
                en: (rows[i][3] || "").toString(), options: opts });
  }
  if (!pool.length) {
    return { status: "noMore", remaining,
      message: "Wszystkie pytania rozwiązane! / 全問クリアです。新しい問題をお待ちください" };
  }
  for (let i = pool.length - 1; i > 0; i--) {          // シャッフル
    const j = Math.floor(Math.random() * (i + 1));
    [pool[i], pool[j]] = [pool[j], pool[i]];
  }
  return { status: "ok", remaining, daily: triviaDaily(), point: triviaPoint(),
           questions: pool.slice(0, remaining) };
}

// ── POST action=submitTrivia {memberId, qid, answer} ─────────────
function handleTriviaAnswer(data) {
  if (!triviaOn()) return jsonResponse({ status: "off" });
  const id  = (data.memberId || "").toString().trim();
  const qid = (data.qid || "").toString().trim();
  const ans = Number(data.answer);
  if (!id || !qid || !ans) return jsonResponse({ status: "error", message: "Brak danych" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const member = findMember(id);
  if (!member) return jsonResponse({ status: "notMember" });
  if (!hasOrderedToday(ss, member.email)) {
    return jsonResponse({ status: "noOrder",
      message: "Quiz dostępny po złożeniu zamówienia / 本日ご注文のお客様が対象です" });
  }

  const sh = ss.getSheetByName(SHEET_TRIVIA);
  if (!sh) return jsonResponse({ status: "error", message: "no sheet" });
  const rows = sh.getDataRange().getValues();
  let q = null;
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0] || "").toString().trim() === qid) { q = rows[i]; break; }
  }
  if (!q) return jsonResponse({ status: "error", message: "no question" });

  // 1日の上限と重複回答をサーバー側で判定（画面の細工では抜けられない）
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  let lg = ss.getSheetByName(SHEET_TRIVIALOG);
  if (!lg) {
    lg = ss.insertSheet(SHEET_TRIVIALOG);
    lg.appendRow(["Data / 日時", "ID członka", "Imię / Name", "Pytanie / 問題",
                  "Odpowiedź", "Wynik / Result", "Punkty"]);
  }
  const myId = (member.memberId || "").toString().trim().toLowerCase();
  const lr = lg.getDataRange().getValues();
  let todayCount = 0;
  for (let i = 1; i < lr.length; i++) {
    if ((lr[i][1] || "").toString().trim().toLowerCase() !== myId) continue;
    const rowQid = (lr[i][3] || "").toString().split(" ")[0];
    if (rowQid === qid) {
      return jsonResponse({ status: "already",
        message: "To pytanie już rozwiązane / この問題は回答済みです" });
    }
    if ((lr[i][0] || "").toString().slice(0, 10) === today) todayCount++;
  }
  if (todayCount >= triviaDaily()) {
    return jsonResponse({ status: "doneToday", remaining: 0,
      message: "Dziś już wszystkie pytania! / 本日の分は終了しました" });
  }

  const correct = (ans === Number(q[8]));
  const pt      = correct ? triviaPoint() : 0;
  let newPt     = member.points || 0;
  let rankUp    = false;

  if (correct) {
    const oldPt = newPt;
    newPt = oldPt + pt;
    rankUp = getRank(oldPt).name !== getRank(newPt).name;
    try { awardMilestoneCoupons(member.email, member.name, oldPt, newPt); }
    catch (e) { Logger.log("trivia coupon: " + e); }
    // ポイントのみ加算（来店回数＝杯数には影響させない）
    updateMember(ss, member.email, { points: newPt, rank: getRank(newPt).name,
      _source: "Quiz o Japonii / 日本クイズ (" + qid + ")" });
  }

  lg.appendRow([Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd HH:mm"),
                member.memberId || "", member.name || "",
                qid + " " + (q[2] || "").toString().slice(0, 60),
                ans, correct ? "⭕ Poprawnie" : "❌ Błędnie", pt]);

  return jsonResponse({
    status: "ok", correct, awarded: pt, newPt, rankUp,
    rank: getRank(newPt),
    remaining: Math.max(0, triviaDaily() - todayCount - 1),
    explainPL: (q[9]  || "").toString(),
    explainEN: (q[10] || "").toString()
  });
}

// ════════════════════════════════════════════════════════════════════
//  ⏰ 時間帯ボーナス（ハッピーアワー）— v6.5
//    閑散時間帯の来店に追加ポイントを付与する販促機能。
//    割引ではなくポイントなので、粗利を減らさず再来店につながる。
//    初期値は OFF。期間（開始日・終了日）も指定できる。
// ════════════════════════════════════════════════════════════════════
function hhGet(k, def) {
  try { return PropertiesService.getScriptProperties().getProperty(k) || def; }
  catch (e) { return def; }
}

// 現在ハッピーアワーが有効か判定し、内容を返す
//  戻り値 { active:boolean, bonus:number, label:string, ... }
function happyHourStatus() {
  const on = hhGet("HH_MODE", "off") === "on";
  const info = {
    active: false, enabled: on,
    bonus: Number(hhGet("HH_BONUS", "10")) || 10,
    from:  Number(hhGet("HH_FROM", "18")),
    to:    Number(hhGet("HH_TO", "21")),
    days:  hhGet("HH_DAYS", "1,2,3,4,5"),          // 1=月〜7=日
    start: hhGet("HH_START", ""), end: hhGet("HH_END", "")
  };
  if (!on) return info;

  const now   = new Date();
  const today = Utilities.formatDate(now, "Europe/Warsaw", "yyyy-MM-dd");
  if (info.start && today < info.start) return info;   // 期間前
  if (info.end   && today > info.end)   return info;   // 期間終了 → 自動で止まる

  const dow  = Utilities.formatDate(now, "Europe/Warsaw", "u");
  const hour = Number(Utilities.formatDate(now, "Europe/Warsaw", "H"));
  if (info.days.split(",").map(x => x.trim()).indexOf(dow) === -1) return info;
  if (hour < info.from || hour >= info.to) return info;

  info.active = true;
  return info;
}

// メニュー：ハッピーアワーの設定（ボタン操作だけで完結）
function setupHappyHour() {
  const ui = SpreadsheetApp.getUi();
  const st = happyHourStatus();
  const dowName = { "1":"Pon", "2":"Wt", "3":"Śr", "4":"Czw", "5":"Pt", "6":"Sob", "7":"Ndz" };
  const daysTxt = st.days.split(",").map(d => dowName[d.trim()] || d).join(", ");

  const cur = "Teraz / 現在: " + (st.enabled ? "ON" : "OFF") +
    "  |  " + daysTxt + "  " + st.from + ":00–" + st.to + ":00" +
    "  |  +" + st.bonus + " pkt" +
    (st.start || st.end ? "  |  " + (st.start || "…") + " → " + (st.end || "…") : "");

  const r = ui.prompt("⏰ Happy Hour / 時間帯ボーナス",
    cur + "\n\n" +
    "Wpisz ustawienia oddzielone przecinkami / 設定をカンマ区切りで入力:\n" +
    "  ON/OFF, dni, od, do, punkty, start, koniec\n\n" +
    "Dni / 曜日: 1=Pon 2=Wt 3=Śr 4=Czw 5=Pt 6=Sob 7=Ndz\n" +
    "  zakres „2-4”, lista „2/3/4/7”, mieszane „2-4/7”\n" +
    "  （範囲は 2-4、飛び飛びは 2/3/4/7、混在は 2-4/7）\n\n" +
    "Przykład / 例:\n" +
    "  ON,2-4/7,18,20,10,2026-11-01,2027-01-31\n" +
    "  → wt–czw i ndz 18:00–20:00, +10 pkt, 1.11.2026–31.01.2027\n" +
    "  （火〜木＋日の18:00〜20:00に+10pt）\n\n" +
    "Aby wyłączyć wpisz po prostu / 止めるには:  OFF",
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const parts = r.getResponseText().split(",").map(x => x.trim());
  const props = PropertiesService.getScriptProperties();

  if ((parts[0] || "").toUpperCase() === "OFF") {
    props.setProperty("HH_MODE", "off");
    ui.alert("✅ Happy Hour: OFF\n時間帯ボーナスを停止しました。");
    return;
  }
  if ((parts[0] || "").toUpperCase() !== "ON") {
    ui.alert("⚠️ Zacznij od ON albo OFF / 最初に ON か OFF を入力してください");
    return;
  }

  // v6.6修正：曜日は「2-4」「2/3/4/7」「2-4/7」の形で指定する。
  // （カンマは項目の区切りに使っているため、曜日側では使えない）
  let days = parts[1] || "1-5";
  const dset = [];
  days.split("/").forEach(seg => {
    seg = seg.trim();
    if (!seg) return;
    if (seg.indexOf("-") !== -1) {
      const [a, b] = seg.split("-").map(Number);
      for (let d = a; d <= b; d++) if (d >= 1 && d <= 7) dset.push(d);
    } else {
      const d = Number(seg);
      if (d >= 1 && d <= 7) dset.push(d);
    }
  });
  if (!dset.length) { ui.alert("⚠️ Błędne dni / 曜日の指定が不正です（例: 2-4/7）"); return; }
  days = dset.sort().join(",");
  const from  = Number(parts[2]);
  const to    = Number(parts[3]);
  const bonus = Number(parts[4]);
  if (!(from >= 0 && from <= 23) || !(to >= 1 && to <= 24) || to <= from) {
    ui.alert("⚠️ Błędne godziny / 時間の指定が不正です（例: 18 と 21）");
    return;
  }
  if (!(bonus > 0 && bonus <= 50)) {
    ui.alert("⚠️ Punkty 1–50 / ポイントは1〜50で指定してください");
    return;
  }

  props.setProperty("HH_MODE", "on");
  props.setProperty("HH_DAYS", days);
  props.setProperty("HH_FROM", String(from));
  props.setProperty("HH_TO", String(to));
  props.setProperty("HH_BONUS", String(bonus));
  props.setProperty("HH_START", parts[5] || "");
  props.setProperty("HH_END", parts[6] || "");
  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}

  ui.alert("✅ Zapisano / 保存しました",
    "Happy Hour: ON\n" +
    "Dni / 曜日: " + days.split(",").map(d => dowName[d] || d).join(", ") + "\n" +
    "Godziny / 時間: " + from + ":00–" + to + ":00\n" +
    "Bonus: +" + bonus + " pkt\n" +
    "Okres / 期間: " + (parts[5] || "od zaraz") + " → " + (parts[6] || "bez końca") + "\n\n" +
    "Poza tym okresem funkcja wyłącza się sama.\n期間外は自動的に停止します。",
    ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════════
//  📊 客足分析（曜日 × 時間帯）— v6.5
//    注文シートを集計し、いつが空いているかを色分けで可視化する。
//    「感覚」ではなく実データで閑散時間帯を決めるためのツール。
// ════════════════════════════════════════════════════════════════════
const SHEET_TRAFFIC = "客足分析 / Analiza ruchu";

function buildTrafficReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(SHEET_ORDERS);
  if (!src) { ui.alert("Brak arkusza zamówień / 注文シートがありません"); return; }

  // 集計期間を聞く（既定90日）
  const resp = ui.prompt("📊 Analiza ruchu / 客足分析",
    "Ile ostatnich dni przeanalizować? / 何日分を集計しますか？\n" +
    "(np. 30, 90, 365 — Enter = 90)",
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const days = Math.max(1, Number(resp.getResponseText().trim()) || 90);

  const data  = src.getDataRange().getValues();
  const limit = new Date();
  limit.setDate(limit.getDate() - days);

  // 注文単位で集計（1注文＝複数行なので注文IDでまとめる）
  const seen  = Object.create(null);
  const cell  = {};        // "曜日|時" → {orders, revenue, guests}
  let totalOrders = 0, totalRev = 0;

  for (let i = 1; i < data.length; i++) {
    const id = (data[i][0] || "").toString().trim();
    const t  = data[i][1];
    if (!id || !(t instanceof Date) || t < limit) continue;
    if (seen[id]) continue;              // 同一注文の2行目以降は数えない
    seen[id] = true;

    const dow  = Number(Utilities.formatDate(t, "Europe/Warsaw", "u"));  // 1=月〜7=日
    const hour = Number(Utilities.formatDate(t, "Europe/Warsaw", "H"));
    const rev  = Number(data[i][18]) || 0;
    const g    = parseInt((data[i][3] || "").toString(), 10) || 0;

    const k = dow + "|" + hour;
    if (!cell[k]) cell[k] = { orders: 0, revenue: 0, guests: 0 };
    cell[k].orders++;
    cell[k].revenue += rev;
    cell[k].guests  += g;
    totalOrders++;
    totalRev += rev;
  }

  if (!totalOrders) {
    ui.alert("Brak danych w tym okresie / この期間のデータがありません");
    return;
  }

  // 出力シートを作り直す
  let sh = ss.getSheetByName(SHEET_TRAFFIC);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_TRAFFIC);

  const dowName = ["", "Pon / 月", "Wt / 火", "Śr / 水", "Czw / 木",
                   "Pt / 金", "Sob / 土", "Ndz / 日"];

  // 実際に注文があった時間帯だけを列にする
  const hours = [];
  for (let h = 0; h < 24; h++) {
    for (let d = 1; d <= 7; d++) if (cell[d + "|" + h]) { hours.push(h); break; }
  }

  sh.appendRow(["📊 Analiza ruchu / 客足分析 — ostatnie " + days + " dni",
    "Zamówienia razem / 総注文数: " + totalOrders,
    "Obrót / 売上: " + Math.round(totalRev) + " zł"]);
  sh.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#2c1810")
    .setFontColor("#c8a96e").setFontSize(11);
  sh.appendRow([]);

  // ── 表1：注文件数 ──
  writeGrid(sh, "🧾 Liczba zamówień / 注文件数", hours, dowName, cell, "orders");
  // ── 表2：売上 ──
  writeGrid(sh, "💰 Obrót (zł) / 売上", hours, dowName, cell, "revenue");

  // ── 空いている時間帯トップ5 ──
  const list = [];
  Object.keys(cell).forEach(k => {
    const [d, h] = k.split("|").map(Number);
    list.push({ d, h, o: cell[k].orders, r: cell[k].revenue });
  });
  list.sort((a, b) => a.o - b.o);

  const row = sh.getLastRow() + 2;
  sh.getRange(row, 1).setValue("🕳 Najspokojniejsze godziny / 最も空いている時間帯")
    .setFontWeight("bold").setFontColor("#c0392b").setFontSize(11);
  sh.getRange(row + 1, 1, 1, 4).setValues([[
    "Dzień / 曜日", "Godzina / 時間帯", "Zamówienia / 件数", "Obrót / 売上 (zł)"]])
    .setFontWeight("bold").setBackground("#fdf3ea");
  list.slice(0, 8).forEach((x, i) => {
    sh.getRange(row + 2 + i, 1, 1, 4).setValues([[
      dowName[x.d], x.h + ":00–" + (x.h + 1) + ":00", x.o, Math.round(x.r)]]);
  });

  sh.setColumnWidth(1, 130);
  for (let c = 2; c <= hours.length + 1; c++) sh.setColumnWidth(c, 78);
  sh.setFrozenColumns(1);

  ui.alert("✅ Gotowe / 完了",
    "Analiza " + days + " dni: " + totalOrders + " zamówień.\n" +
    "Zobacz arkusz „" + SHEET_TRAFFIC + "”.\n\n" +
    "Czerwone pola = najspokojniej（赤いセルほど空いている時間帯）",
    ui.ButtonSet.OK);
}

// 曜日×時間の表を1つ書き出す（値が小さいほど赤＝空いている）
function writeGrid(sh, title, hours, dowName, cell, field) {
  const top = sh.getLastRow() + 2;
  sh.getRange(top, 1).setValue(title).setFontWeight("bold").setFontSize(11)
    .setFontColor("#c0392b");

  const head = ["Dzień \\ Godz."].concat(hours.map(h => h + ":00"));
  sh.getRange(top + 1, 1, 1, head.length).setValues([head])
    .setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e")
    .setHorizontalAlignment("center").setFontSize(9);

  let max = 0;
  for (let d = 1; d <= 7; d++) hours.forEach(h => {
    const c = cell[d + "|" + h];
    if (c && c[field] > max) max = c[field];
  });

  for (let d = 1; d <= 7; d++) {
    const vals = hours.map(h => {
      const c = cell[d + "|" + h];
      return c ? (field === "revenue" ? Math.round(c[field]) : c[field]) : 0;
    });
    const r = top + 1 + d;
    sh.getRange(r, 1).setValue(dowName[d]).setFontWeight("bold").setBackground("#fdf3ea");
    sh.getRange(r, 2, 1, vals.length).setValues([vals])
      .setHorizontalAlignment("center");
    // 値に応じた色（少ない=赤 / 多い=緑）
    vals.forEach((v, i) => {
      const ratio = max ? v / max : 0;
      const bg = v === 0 ? "#f5f5f5"
        : ratio < 0.25 ? "#f8d7da"
        : ratio < 0.5  ? "#fde8c8"
        : ratio < 0.75 ? "#e8f5c8" : "#cdebc5";
      sh.getRange(r, 2 + i).setBackground(bg);
    });
  }
}

// ════════════════════════════════════════════════════════════════════
//  ⭐ おすすめ（期間限定・値引き商品）— v6.0
//    シートを編集するだけで注文アプリの先頭に「⭐おすすめ」タブが出る。
//    掲載期間を過ぎた行は自動で消えるので、消し忘れ事故が起きない。
// ════════════════════════════════════════════════════════════════════
const SHEET_PROMO = "おすすめ / Polecane";

function setupPromoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let sh = ss.getSheetByName(SHEET_PROMO);
  if (sh) {
    const r = ui.alert("おすすめシートは既に存在します",
      "作り直しますか？（現在の内容は消えます）", ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet(SHEET_PROMO, 1);
  const headers = [
    "商品ID / ID（メニューシートのB列）",  // A
    "開始日 / Od (yyyy-mm-dd)",           // B
    "終了日 / Do (yyyy-mm-dd)",           // C
    "特価 zł（空欄=通常価格のまま）",       // D
    "バッジPL / Badge PL",                // E
    "バッジJP / バッジ日本語",             // F
    "表示順 / Kolejność",                 // G
  ];
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold")
    .setBackground("#c8a96e").setFontColor("#2c1810")
    .setHorizontalAlignment("center").setFontSize(10);
  [260, 150, 150, 170, 190, 170, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);

  // B列・C列（開始日・終了日）に日付ピッカーを500行分設定。
  // これでどの行でもダブルクリック→カレンダーから選べる（数字を手入力しなくてよい）。
  applyPromoDatePicker_(sh);

  // 記入例（グレー斜体・そのまま消してOK）
  sh.appendRow(["r4", "2026-08-01", "2026-08-31", 39,
                "Oferta sierpniowa −8 zł", "8月のおすすめ −8zł", 1]);
  sh.getRange(2, 1, 1, 7).setFontColor("#999999").setFontStyle("italic");
  sh.getRange(2, 1).setNote("これは記入例です。実際に使わないなら、行番号2を右クリック→「行を削除」で消してください。");

  // 右側に「書き方」凡例を常設（記入例を消しても残る）
  writePromoLegend_(sh);

  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}
  ui.alert("✅ おすすめシートを作成しました",
    "使い方はシート右側の「📖 書き方」に常時表示しています。\n\n" +
    "・A列に商品ID（メニューシートのB列の値、例 r4）\n" +
    "・B/C列に掲載期間（期間外は自動で非表示）\n" +
    "・D列に特価（空欄なら通常価格で「おすすめ」表示のみ）\n" +
    "・E/F列のバッジは商品名の下に表示されます\n\n" +
    "反映は最大1分（メニューキャッシュ）です。", ui.ButtonSet.OK);
}

// ── おすすめシート右側の「書き方」凡例を書き込む（常設・再実行で更新） ──
//   既存の運用中シートにも後付けできるよう関数として独立させている。
function writePromoLegend_(sh) {
  const C = 9;  // I列から（I=見出し, J=PL, K=JP）
  const lines = [
    ["📖 Jak dodać „Polecane” / おすすめの追加方法", "", ""],
    ["", "", ""],
    ["Kolumna / 列", "Co wpisać (PL)", "説明（日本語）"],
    ["A", "ID dania z arkusza „Menu” (kol. B). Np. r4, o1, a4",
          "商品ID（メニューシートB列）。例 r4, o1, a4"],
    ["B", "Data OD (rrrr-mm-dd). Puste = od dziś",
          "開始日。空欄なら即日から"],
    ["C", "Data DO (rrrr-mm-dd). Po tej dacie znika sam",
          "終了日。過ぎると自動で消える"],
    ["D", "Cena promocyjna (zł). Puste = cena normalna",
          "特価。空欄なら通常価格のまま"],
    ["E", "Znaczek po polsku. Np. Oferta −8 zł",
          "バッジPL（商品名の下に表示）"],
    ["F", "Znaczek po japońsku. Np. 8月のおすすめ",
          "バッジ日本語"],
    ["G", "Kolejność (1, 2, 3…). Mniejsze = wyżej",
          "表示順。小さい数字が先頭"],
    ["", "", ""],
    ["🗑 Koniec", "Usuń wiersz LUB niech minie data w kol. C",
                  "行を削除、または終了日を過ぎさせる"],
    ["⏱ Kiedy", "Zmiany widać po ok. 1 minucie",
                 "反映は最大1分"],
    ["📷 Zdjęcie", "Plik images/[ID].jpg — wspólny z aplikacją i stroną",
                    "写真は images/[商品ID].jpg（アプリ・HP共通）"],
    ["⚠️ Uwaga", "Nieistniejące ID jest pomijane",
                  "存在しない商品IDは無視される"],
  ];
  sh.getRange(1, C, lines.length, 3).setValues(lines);

  // 見出し（I1:K1を結合）
  sh.getRange(1, C, 1, 3).merge().setBackground("#2c1810").setFontColor("#c8a96e")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  // 列見出し行（3行目）
  sh.getRange(3, C, 1, 3).setBackground("#c8a96e").setFontColor("#2c1810")
    .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  // 本文
  sh.getRange(4, C, lines.length - 3, 3).setFontSize(9)
    .setVerticalAlignment("top").setWrap(true);
  sh.getRange(1, C, lines.length, 1).setFontWeight("bold").setHorizontalAlignment("center");
  sh.getRange(1, C, lines.length, 3).setBackground("#fdf3ea");
  sh.getRange(1, C, 1, 3).setBackground("#2c1810");     // 見出し帯を上書きで濃色に
  sh.getRange(3, C, 1, 3).setBackground("#c8a96e");     // 列見出しを金色に戻す

  sh.setColumnWidth(C, 78);       // I 列名
  sh.setColumnWidth(C + 1, 320);  // J ポーランド語
  sh.setColumnWidth(C + 2, 220);  // K 日本語
  sh.setColumnWidth(C - 1, 24);   // H 区切り
}

// ── おすすめシートB・C列に日付ピッカーを設定（500行分） ─────────
//   ダブルクリックでカレンダーが出て日付を選べるようにする。
//   数字のyyyy-mm-dd手入力より速く、書式ミス（typo）も防げる。
function applyPromoDatePicker_(sh) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText("カレンダーから日付を選んでください / Wybierz datę z kalendarza")
    .build();
  sh.getRange(2, 2, 500, 2).setDataValidation(rule);   // B2:C501
  sh.getRange(2, 2, 500, 2).setNumberFormat("yyyy-mm-dd");
}

// ── 既存のおすすめシートに日付ピッカーだけ後付け（作り直さない） ──
//   メニュー「⭐ 日付選択を有効化」から実行。運用中でも安全。
function enablePromoDatePicker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(SHEET_PROMO);
  if (!sh) {
    ui.alert("おすすめシートがありません", "先に「⭐ おすすめシートを作成」を実行してください。", ui.ButtonSet.OK);
    return;
  }
  applyPromoDatePicker_(sh);
  ui.alert("✅ 日付選択を有効化しました",
    "B列（開始日）・C列（終了日）は、どの行でもセルをダブルクリックすると\n" +
    "カレンダーが出て日付を選べます。手入力も引き続き可能です。\n\n" +
    "対象：2〜501行目", ui.ButtonSet.OK);
}

// ── 既存のおすすめシートに凡例だけ後付け（作り直さない） ──
//   メニュー「⭐ おすすめの書き方を表示」から実行。運用中でも安全。
function showPromoLegend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let sh = ss.getSheetByName(SHEET_PROMO);
  if (!sh) {
    const r = ui.alert("おすすめシートがありません",
      "先に作成しますか？（setupPromoSheet を実行します）", ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) setupPromoSheet();
    return;
  }
  writePromoLegend_(sh);
  ui.alert("✅ 「書き方」をシート右側（I列〜）に表示しました",
    "この凡例は記入データを消しても残ります。\n" +
    "いつでもここで記入方法を確認できます。", ui.ButtonSet.OK);
}

// 掲載中のおすすめを1カテゴリにまとめて返す（該当なしならnull）
function buildPromoCategory(ss, menu) {
  const sh = ss.getSheetByName(SHEET_PROMO);
  if (!sh) return null;

  // 全商品を id で引けるようにする
  const byId = Object.create(null);
  menu.forEach(c => c.items.forEach(it => { byId[it.id] = it; }));

  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const rows  = sh.getDataRange().getValues();
  const out   = [];

  for (let i = 1; i < rows.length; i++) {
    const r  = rows[i];
    const id = (r[0] || "").toString().trim();
    if (!id || !byId[id]) continue;                    // 存在しないIDは無視
    const from = fmtPromoDate(r[1]);
    const to   = fmtPromoDate(r[2]);
    if (from && today < from) continue;                // 掲載前
    if (to   && today > to)   continue;                // 掲載終了 → 自動で消える

    const src   = byId[id];
    const price = Number(r[3]) > 0 ? Number(r[3]) : src.price;
    const item  = JSON.parse(JSON.stringify(src));     // 元商品を壊さないよう複製

    // ID衝突を避ける（同じ商品を通常カテゴリからも注文できるため）
    item.id     = "p_" + src.id;
    item.imgKey = src.imgKey || src.id;
    item.price  = price;
    if (price < src.price) item.origPrice = src.price;  // 取り消し線表示用
    item.promoPL = (r[4] || "").toString();
    item.promoJP = (r[5] || "").toString();
    if (to) item.promoUntil = to;
    item.promoOrder = Number(r[6]) || 99;
    out.push(item);
  }
  if (!out.length) return null;

  out.sort((a, b) => a.promoOrder - b.promoOrder);
  return {
    id: "osusume",
    namePL: "⭐ Polecane",
    nameEN: "⭐ Weekly & Monthly Picks",
    subcategorized: false,
    items: out
  };
}

// 日付セルが Date でも文字列でも yyyy-MM-dd に揃える
function fmtPromoDate(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Warsaw", "yyyy-MM-dd");
  return v.toString().trim().slice(0, 10);
}

const SHEET_HIDDEN   = "商品表示管理 / Item Visibility";

// ── 商品表示管理シート（なければ作成） ───────────────────────────
function getHiddenSheet(ss) {
  let sh = ss.getSheetByName(SHEET_HIDDEN);
  if (!sh) {
    sh = ss.insertSheet(SHEET_HIDDEN);
    sh.appendRow(["非表示商品ID / Hidden item IDs", "更新日時 / Updated"]);
    sh.getRange(1,1,1,2).setFontWeight("bold").setBackground("#2c1810")
      .setFontColor("#c8a96e").setHorizontalAlignment("center");
    sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 180);
    sh.getRange(2,2).setValue("※ このシートはitems.htmlから自動管理されます。手動編集不要。");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── 非表示IDリスト取得 ────────────────────────────────────────────
function getHiddenItemIds(ss) {
  const sh = getHiddenSheet(ss);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getValues()
    .map(r => (r[0]||"").toString().trim()).filter(Boolean);
}

// ── 非表示IDリスト保存（要パスワード） ────────────────────────────
function handleSetHiddenItems(data) {
  const auth4 = verifyStaffPassword(data.password);
  if (auth4 !== "ok") return jsonResponse({ status:"error", message: STAFF_AUTH_MSG[auth4] });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getHiddenSheet(ss);
  const hidden = Array.isArray(data.hidden) ? data.hidden.map(String) : [];

  // 既存クリア → 書き込み
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  const now = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd HH:mm:ss");
  if (hidden.length) {
    sh.getRange(2, 1, hidden.length, 1).setValues(hidden.map(id => [id]));
    sh.getRange(2, 2).setValue(now);
  } else {
    sh.getRange(2, 2).setValue(now + "（全商品表示中）");
  }
  return jsonResponse({ status:"ok", count: hidden.length,
    message: "保存完了 / Zapisano: " + hidden.length + "品を非表示" });
}

// ── ポイント履歴シート取得（なければ自動作成） ──────────────────
function getPointLogSheet(ss) {
  let sh = ss.getSheetByName(SHEET_POINTLOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_POINTLOG);
    const headers = [
      "Data / Timestamp",          // A
      "🪪 ID członka / Member ID", // B
      "E-mail / Kontakt",          // C
      "Imię / Name",               // D
      "Zmiana / Change",           // E
      "Razem / Total",             // F
      "Ranga / Rank",              // G
      "Źródło / Source"            // H
    ];
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold").setBackground("#2c1810")
      .setFontColor("#c8a96e").setHorizontalAlignment("center").setFontSize(10);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 130);
    sh.setColumnWidth(3, 200); sh.setColumnWidth(4, 140);
    sh.setColumnWidth(5, 80);  sh.setColumnWidth(6, 80);
    sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 220);
    // E列は「+10pt」等を入れるため、列全体をテキスト書式に固定（数式誤解釈の防止）
    sh.getRange("E2:E").setNumberFormat("@");
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── 既存のポイント履歴シートを修復（#ERROR! 表示を直す・1回実行でOK）──
//  E列を数式扱いしていたため #ERROR! になった行を、正しい表記に直す。
function fixPointLogFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_POINTLOG);
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert("ポイント履歴シートがありません"); return; }

  const last = sh.getLastRow();
  if (last < 2) { ui.alert("修復対象の行がありません"); return; }

  const rng     = sh.getRange(2, 5, last - 1, 1);
  const totals  = sh.getRange(2, 6, last - 1, 1).getValues();  // F列「累計」
  const formulas = rng.getFormulas();
  const values   = rng.getValues();
  const out = [];
  let fixed = 0;

  for (let i = 0; i < values.length; i++) {
    const f = (formulas[i][0] || "").toString();
    const v = (values[i][0] || "").toString();
    if (f) {
      // 「=+10pt」のような誤変換 → 元の「+10pt」に復元
      out.push([f.replace(/^=/, "")]);
      fixed++;
    } else if (v === "" || v.indexOf("#") === 0) {
      out.push(["—"]);   // 復元不能な行は記号で埋める（累計はF列に残っている）
      fixed++;
    } else {
      out.push([v]);
    }
  }
  rng.setNumberFormat("@").setValues(out);
  sh.getRange("E2:E").setNumberFormat("@");   // 以後の行も数式化しないよう固定

  ui.alert("✅ 修復完了",
    fixed + " 行を修正しました。\n以後は自動的にテキストとして記録されます。\n" +
    "※累計ポイント（F列）と会員シートの値は元から正常です。", ui.ButtonSet.OK);
}

// ── ポイント履歴を1行記録 ─────────────────────────────────────────
function logPointHistory(ss, memberId, email, name, oldPt, newPt, rankName, rankIcon, source) {
  try {
    const sh    = getPointLogSheet(ss);
    const delta = newPt - oldPt;
    const ts    = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd HH:mm:ss");
    sh.appendRow([
      ts, memberId || "", email || "", name || "",
      (delta >= 0 ? "+" : "") + delta + "pt",
      newPt + "pt",
      (rankIcon || "") + " " + (rankName || ""),
      source || "その他 / Other"
    ]);
    const row = sh.getLastRow();
    // v6.1修正：「+10pt」はGoogleスプレッドシートに数式と誤解され #ERROR! になるため、
    // E列を先にテキスト書式へ変更してから値を入れ直す。
    sh.getRange(row, 5).setNumberFormat("@").setValue((delta >= 0 ? "+" : "") + delta + "pt");
    // 増減に応じて色分け（付与=緑・取消/失効=赤）
    sh.getRange(row, 5).setFontWeight("bold")
      .setFontColor(delta >= 0 ? "#2d7a3a" : "#c0392b")
      .setHorizontalAlignment("center");
    sh.getRange(row, 6).setHorizontalAlignment("center").setFontWeight("bold");
  } catch(e) {
    Logger.log("Point log error: " + e.toString());
  }
}

// ── 係員シートのセットアップ（初回1回実行・列構成変更時も再実行） ──
//  v5.3：ボタンをチェックボックス化（クリックで確実に発火）
//  列構成：A Email / B 氏名 / C 現在pt / D ランク /
//          E ☑＋10pt（🍜ラーメン1杯） / F ☑＋5pt（🎌クイズ正解） /
//          G ☑↩️取消 / H 次のランクまで / I 操作ログ
function setupStaffSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let sh = ss.getSheetByName(SHEET_STAFF);
  if (sh) {
    const r = ui.alert("Odbuduj arkusz obsługi / Rebuild staff sheet",
      "Arkusz „" + SHEET_STAFF + "” zostanie usunięty i utworzony na nowo.\n" +
      "Historia z kolumny I zniknie, ale punkty członków pozostaną nienaruszone.\n\n" +
      "Kontynuować? / Continue?（係員シートを作り直します）",
      ui.ButtonSet.OK_CANCEL);
    if (r !== ui.Button.OK) return;
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet(SHEET_STAFF);

  const headers = [
    "E-mail / Kontakt",              // A
    "Imię i nazwisko / Name",        // B
    "Punkty / Points",               // C
    "Ranga / Rank",                  // D
    "🍜 +10pt (ramen)",              // E ← ラーメン1杯（チェックで付与）
    "🎌 +5pt (quiz)",                // F ← 日本語クイズ正解（チェックで付与）
    "↩️ Cofnij / Undo",              // G ← チェックで取消（額は入力）
    "Do następnej rangi / To next",  // H
    "Historia / Log",                // I
  ];
  sh.appendRow(headers);

  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e")
    .setHorizontalAlignment("center").setFontSize(10);
  sh.getRange(1, 5).setBackground("#2d7a3a").setFontColor("#ffffff");   // +10
  sh.getRange(1, 6).setBackground("#8e44ad").setFontColor("#ffffff");   // +5
  sh.getRange(1, 7).setBackground("#e74c3c").setFontColor("#ffffff");   // 取消

  // チェックボックス（500行分。会員が増えても対応）
  sh.getRange(2, 5, 500, 3).insertCheckboxes();

  [220, 140, 80, 110, 80, 80, 80, 160, 300].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);

  syncStaffSheet();
  Logger.log("✅ 係員シート作成完了 / Staff sheet created (checkbox v5.3)");
}

// ── 会員データを係員シートに同期 ──────────────────────────────────
function syncStaffSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = getMemberSheet(ss);
  const staff   = ss.getSheetByName(SHEET_STAFF);
  if (!staff) { Logger.log("係員シートがありません。setupStaffSheet()を実行してください。"); return; }

  const mData = members.getDataRange().getValues();
  const RANK_COLORS = {
    Sakura:"#fce4ec", Ume:"#fbe9e7",
    Take:"#e8f5e9", Matsu:"#fff8e1", Fuji:"#e3f2fd"
  };

  // 既存データを削除（ヘッダー以外。ログI列も同期のたびにリセットされる点に注意）
  const lastRow = staff.getLastRow();
  if (lastRow > 1) staff.getRange(2, 1, lastRow - 1, 9).clearContent();

  let out = 2;  // 書き込み先行（空メール行を詰める）
  for (let i = 1; i < mData.length; i++) {
    // 会員シート: A=会員ID B=Email C=氏名 K=累計pt L=ランク N=次ランク O=最終来店
    const email     = mData[i][1] || "";
    const name      = mData[i][2] || "";
    const points    = Number(mData[i][10]) || 0;
    const rankName  = mData[i][11] || "Sakura";
    const toNext    = mData[i][13] || "";
    if (!email) continue;

    const rankInfo = getRank(points);
    staff.getRange(out, 1).setValue(email);
    staff.getRange(out, 2).setValue(name);
    staff.getRange(out, 3).setValue(points).setHorizontalAlignment("center").setFontWeight("bold");
    staff.getRange(out, 4).setValue(rankInfo.icon + " " + rankName).setHorizontalAlignment("center");
    staff.getRange(out, 5, 1, 3).setValues([[false, false, false]]);  // チェックボックスをOFFに
    staff.getRange(out, 8).setValue(toNext);

    const bgColor = RANK_COLORS[rankName] || "#ffffff";
    staff.getRange(out, 1, 1, 4).setBackground(bgColor);
    staff.getRange(out, 8, 1, 2).setBackground(bgColor);
    out++;
  }

  Logger.log("✅ 係員シート同期完了 / Staff sheet synced: " + (out - 2) + " members");
}

// ── セルクリック検知：チェックボックスをONにするとポイント操作 ────
//  v5.3で修正：旧方式（「＋10pt」の文字セル）はクリックではonEditが
//  発火せず操作不能だった。チェックボックスはクリック＝編集なので確実。
function onEdit(e) {
  const sh    = e.source.getActiveSheet();
  const range = e.range;
  // ── 書籍シート：返却処理（F列を「在庫あり」に戻すと貸出情報を自動クリア） ──
  if (sh.getName() === SHEET_BOOKS) {
    if (range.getColumn() === 6 && range.getRow() > 1 &&
        (range.getValue()||"").toString() === "在庫あり") {
      sh.getRange(range.getRow(), 7, 1, 4).clearContent();  // 会員ID・名前・貸出日・返却予定日
      try { CacheService.getScriptCache().remove("books_v1"); } catch(e){}
    }
    return;
  }
  if (sh.getName() !== SHEET_STAFF) return;
  if (range.getRow() <= 1) return;
  if (range.getValue() !== true) return;  // チェックONの瞬間のみ反応

  const col = range.getColumn();
  const row = range.getRow();
  range.setValue(false);  // 先にOFFへ戻す（連打・二重発火防止）

  if (col === 5) { confirmAndAddPoint(sh, row, 10); return; }  // 🍜 ラーメン
  if (col === 6) { confirmAndAddPoint(sh, row, 5);  return; }  // 🎌 クイズ正解
  if (col === 7) { confirmAndRemovePoint(sh, row);  return; }  // ↩️ 取消
}

// ── ポイント付与（＋10pt=ラーメン1杯 / ＋5pt=クイズ正解）─────────
//  ＋10pt：来店履歴(history)にも記録（スタンプ1杯分としてカウント）
//  ＋5pt ：ポイントのみ加算。杯数・スタンプグリッドには影響しない
//  v5.3修正：クーポン付与を確認「OK」の後に移動
//  （旧版は確認前に付与していたため、キャンセルしてもクーポンが残った）
function confirmAndAddPoint(sh, row, pts) {
  const email = sh.getRange(row, 1).getValue().toString().trim().toLowerCase();
  if (!email) return;
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const member = findMember(email);
  const ui     = SpreadsheetApp.getUi();
  if (!member) {
    ui.alert("❌ Nie znaleziono członka / Member not found\n" + email);
    return;
  }

  const isQuiz  = (pts === 5);
  const oldPt   = member.points || 0;
  const newPt   = oldPt + pts;
  const oldRank = getRank(oldPt);
  const newRank = getRank(newPt);
  const rankUp  = oldRank.name !== newRank.name;
  // ボーナス到達の見込み（まだ付与しない）
  const nextMs    = getNextMilestone(oldPt);
  const willBonus = nextMs !== null && newPt >= nextMs;

  let msg = "👤 " + (member.name || email) + "\n📧 " + email + "\n";
  msg += "━━━━━━━━━━━━━━━━━━\n";
  msg += (isQuiz ? "🎌 Quiz — poprawna odpowiedź: +5pt\n"
                 : "🍜 1 miska ramenu: +10pt\n");
  msg += "Teraz / Now:  " + oldPt + "pt " + oldRank.icon + " " + oldRank.name + "\n";
  msg += "Po / After:   " + newPt + "pt " + newRank.icon + " " + newRank.name + "\n";
  if (rankUp)    msg += "⬆️ Awans rangi! / Rank up!\n";
  if (willBonus) msg += "🎉 Bonus osiągnięty! / Bonus reached!\n";
  msg += "━━━━━━━━━━━━━━━━━━\nOK = dodaj +" + pts + "pt / add +" + pts + "pt";

  if (ui.alert("✅ Potwierdź / Confirm — 付与確認", msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  // ここで初めてクーポン判定・付与（メール送信失敗でもポイント付与は続行）
  let bonus = null;
  try { bonus = awardMilestoneCoupons(email, member.name, oldPt, newPt); }
  catch (err) { Logger.log("staff addPoint coupon: " + err); }

  const today   = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const history = member.history || [];
  const upd = { points:newPt, lastVisit:today, rank:newRank.name,
    _source: isQuiz ? "Arkusz obsługi — quiz +5 / Staff sheet quiz"
                    : "Arkusz obsługi — ramen +10 / Staff sheet ramen" };
  if (!isQuiz) { history.push(today); upd.history = history; }  // 杯数は+10ptのみ加算
  updateMember(ss, email, upd);

  refreshStaffRow(sh, row, newPt, newRank);
  const logTime = Utilities.formatDate(new Date(), "Europe/Warsaw", "MM-dd HH:mm");
  appendStaffLog(sh, row, (isQuiz ? "🎌 " : "🍜 ") + logTime + " +" + pts + "pt " +
    oldPt + "→" + newPt + "pt" +
    (rankUp ? " ⬆️" + newRank.name : "") + (bonus ? " 🎉" : ""));

  ui.alert("✅ Dodano / Added\n" + (member.name || email) + "\n+" + pts + "pt → razem " + newPt + "pt\n" +
    newRank.icon + " " + newRank.name + (rankUp ? "\n⬆️ Awans rangi! / Rank up!" : ""));
}

// ── ↩️取り消し（-10pt / -5pt を選択）────────────────────────────
//  -10pt：当日付与分がある場合のみ。来店履歴からも当日分を1件削除
//  -5pt ：クイズ分の誤付与用。履歴は触らずポイントのみ減算
function confirmAndRemovePoint(sh, row) {
  const email = sh.getRange(row, 1).getValue().toString().trim().toLowerCase();
  if (!email) return;
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const member = findMember(email);
  const ui     = SpreadsheetApp.getUi();
  if (!member) { ui.alert("❌ Nie znaleziono członka / Member not found"); return; }

  // 取り消す額を入力（10 or 5）
  const resp = ui.prompt("↩️ Cofnij punkty / Undo points",
    "👤 " + (member.name || email) + "  (teraz " + (member.points||0) + "pt)\n\n" +
    "Ile punktów cofnąć? Wpisz 10 lub 5:\n" +
    "  10 … 🍜 ramen (tylko z dzisiaj / only today)\n" +
    "   5 … 🎌 quiz (tylko punkty / points only)",
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const amt = parseInt(resp.getResponseText().trim(), 10);
  if (amt !== 10 && amt !== 5) { ui.alert("⚠️ Wpisz 10 albo 5 / Enter 10 or 5"); return; }

  const today      = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
  const history    = member.history || [];
  const todayCount = history.filter(d => d === today).length;
  const oldPt      = member.points || 0;

  if (oldPt < amt) {
    ui.alert("⚠️ Nie można cofnąć / Cannot undo\nCzłonek ma tylko " + oldPt + "pt.");
    return;
  }
  if (amt === 10 && todayCount === 0) {
    ui.alert("⚠️ Nie można cofnąć / Cannot undo\nBrak punktów 🍜 przyznanych dzisiaj.\n(-10pt można cofnąć tylko tego samego dnia)");
    return;
  }

  const newPt    = Math.max(0, oldPt - amt);
  const oldRank  = getRank(oldPt);
  const newRank  = getRank(newPt);
  const rankDown = oldRank.name !== newRank.name;

  let msg = "👤 " + (member.name || email) + "\n";
  msg += (amt === 10 ? "🍜 Ramen: -10pt (usuwa też 1 wizytę)\n"
                     : "🎌 Quiz: -5pt (tylko punkty)\n");
  msg += "Teraz / Now: " + oldPt + "pt " + oldRank.icon + " " + oldRank.name + "\n";
  msg += "Po / After:  " + newPt + "pt " + newRank.icon + " " + newRank.name + "\n";
  if (rankDown) msg += "⬇️ Spadek rangi / Rank down\n";
  msg += "OK = cofnij -" + amt + "pt";

  if (ui.alert("↩️ Potwierdź / Confirm — 取消確認", msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const upd = { points:newPt, rank:newRank.name,
    _source: amt === 10 ? "Arkusz obsługi — cofnięcie -10 / Staff undo ramen"
                        : "Arkusz obsługi — cofnięcie -5 / Staff undo quiz" };
  if (amt === 10) {
    const newHistory   = [...history];
    const lastTodayIdx = newHistory.lastIndexOf(today);
    if (lastTodayIdx !== -1) newHistory.splice(lastTodayIdx, 1);
    upd.history   = newHistory;
    upd.lastVisit = newHistory.length ? newHistory[newHistory.length-1] : "";
  }
  updateMember(ss, email, upd);

  refreshStaffRow(sh, row, newPt, newRank);
  const logTime = Utilities.formatDate(new Date(), "Europe/Warsaw", "MM-dd HH:mm");
  appendStaffLog(sh, row, "↩️ " + logTime + " 取消-" + amt + " " + oldPt + "→" + newPt + "pt" +
    (rankDown ? " ⬇️" + newRank.name : ""));

  ui.alert("↩️ Cofnięto / Undone\n-" + amt + "pt → razem " + newPt + "pt" + (rankDown ? "\n⬇️ Spadek rangi / Rank down" : ""));
}

// ── 係員シート行の表示更新 ────────────────────────────────────────
function refreshStaffRow(sh, row, newPt, newRank) {
  const RANK_COLORS = {Sakura:"#fce4ec",Ume:"#fbe9e7",Take:"#e8f5e9",Matsu:"#fff8e1",Fuji:"#e3f2fd"};
  const bg = RANK_COLORS[newRank.name] || "#ffffff";
  const nextMs = getNextMilestone(newPt);
  const toNext = nextMs !== null ? (nextMs - newPt) + "pt → " + getNextRankName(newPt) : "🎖️ MAX";
  sh.getRange(row, 3).setValue(newPt).setHorizontalAlignment("center").setFontWeight("bold");
  sh.getRange(row, 4).setValue(newRank.icon + " " + newRank.name).setHorizontalAlignment("center");
  sh.getRange(row, 8).setValue(toNext);
  sh.getRange(row, 1, 1, 4).setBackground(bg);
  sh.getRange(row, 8, 1, 2).setBackground(bg);
}

// ── 係員シートI列へログ追記 ──────────────────────────────────────
function appendStaffLog(sh, row, entry) {
  const cell = sh.getRange(row, 9);
  const cur  = cell.getValue().toString();
  cell.setValue(cur ? cur + "\n" + entry : entry)
      .setWrap(true).setVerticalAlignment("top").setFontSize(9);
}
// ── 手動同期ボタン用（メニューから実行可） ──────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🍜 Wabi Navi")
    // ── CODZIENNE / 毎日つかう ──
    .addItem("📅 Rezerwacje: odśwież dziś / 本日の予約を更新", "updateTodaySheet")
    .addItem("🔄 Personel: odśwież / 係員シート更新", "syncStaffSheet")
    .addItem("⭐ Polecane: jak dodać? / おすすめの書き方", "showPromoLegend")
    .addItem("📅 Polecane: kalendarz dat / 日付選択を有効化", "enablePromoDatePicker")

    // ── USTAWIENIA / 設定（切替・オプション） ──
    .addSubMenu(ui.createMenu("⚙️ Ustawienia / 設定")
      .addItem("⏰ Happy Hour / 時間帯ボーナス", "setupHappyHour")
      .addItem("🗾 Quiz o Japonii / 日本クイズ設定", "setupTriviaSettings")
      .addItem("🎚 Lekcje ON/OFF / 日本語講座 公開設定", "toggleLessonSettings")
      .addItem("💬 WhatsApp / 会員WhatsApp設定", "setupWhatsApp")
      .addSeparator()
      .addItem("📸 Zdjęcia: konfiguracja / 写真アップロード設定", "setupPhotoUpload")
      .addItem("📸 Zdjęcia: włącz odbiór / 自動受信をON", "setupPhotoTrigger")
      .addItem("📸 Zdjęcia: wyłącz / 自動受信をOFF", "stopPhotoTrigger"))

    // ── RAPORTY / レポート ──
    .addSubMenu(ui.createMenu("📊 Raporty / 分析")
      .addItem("📊 Analiza ruchu / 客足分析", "buildTrafficReport")
      .addItem("📷 Sprawdź zdjęcia menu / 写真掲載チェック", "checkMenuPhotos")
      .addItem("⏳ Sprawdź wygasłe punkty / 失効チェック", "checkExpiryAndNotify"))

    // ── PIERWSZE URUCHOMIENIE / 初回セットアップ（1回だけ） ──
    .addSubMenu(ui.createMenu("🆕 Pierwsze uruchomienie / 初回設定")
      .addItem("Menu / メニューシート", "setupMenuSheet")
      .addItem("Polecane / おすすめシート", "setupPromoSheet")
      .addItem("Wydarzenia / イベントシート", "setupEventsSheet")
      .addItem("Książki / 書籍シート", "setupBooksSheet")
      .addItem("Lekcje / 日本語講座シート", "setupLessonSheets")
      .addItem("Quiz / クイズシート", "setupTriviaSheet")
      .addItem("Personel / 係員シート", "setupStaffSheet")
      .addSeparator()
      .addItem("Rezerwacje: arkusz / 予約シート", "setupReservationsSheet")
      .addItem("Rezerwacje: trigger formularza / フォーム連携", "setupReservationFormTrigger")
      .addItem("Rezerwacje: trigger „Dziś” / 本日更新の自動化", "setupTodaySheetTrigger")
      .addItem("Punkty: kolumna kuponów / クーポン列", "setupCouponsColumn")
      .addItem("Punkty: codzienny trigger / 毎日の失効チェック", "setupDailyTrigger"))

    // ── KONSERWACJA / 保守・修復（困った時だけ） ──
    .addSubMenu(ui.createMenu("🩹 Konserwacja / 保守・修復")
      .addItem("🩹 Napraw historię wizyt / 来店履歴を修復", "repairMemberHistoryColumn")
      .addItem("🩹 Odtwórz wizyty z logu / 来店日を復元", "repairMemberHistoryFromLog")
      .addItem("🔧 Napraw log punktów / ポイント履歴を修復", "fixPointLogFormat")
      .addItem("🍳 Kolumna kuchni / キッチン用T列を追加", "setupKitchenColumn")
      .addItem("🍳 Przygaś wydane / 提供済みを薄く表示", "dimServedOrderRows")
      .addItem("🛠 Odbuduj personel / 係員シート再作成", "setupStaffSheet"))
    .addToUi();
}



// ════════════════════════════════════════════════════════════════════
//  予約モジュール / Reservations Module（v3で統合済み）
//  doPost / doGet への分岐は上部に反映済み。手動編集は不要です。
//
//  【初回セットアップ】（Apps Scriptエディタで1回ずつ実行）
//    1. Googleフォームの「回答」→ 緑のスプレッドシートアイコン →
//       「既存のスプレッドシートを選択」→ Wabi Naviのシートを選択
//    2. setupReservationsSheet() を実行
//    3. setupReservationFormTrigger() を実行
//    4. ウェブアプリを「新しいバージョン」で再デプロイ
// ════════════════════════════════════════════════════════════════════

const SHEET_RESERVATIONS = "予約 / Rezerwacje";

// ── 予約の運用設定（店側で自由に変更可） ──────────────────────────
const RES_CONFIG = {
  openTime:  "13:00",      // 最初の予約枠
  lastTime:  "21:30",      // 最後の予約枠
  lateFrom:  "20:00",      // この時刻以降の予約が可能な曜日を制限
  lateDays:  [5, 6],       // 0=日,1=月,...5=金,6=土 → 20:00以降は金・土のみ
  closedDays: [],          // 定休日 例: [1] なら月曜休み
  maxGuests: 7,            // オンライン予約の最大人数（8名以上は電話）
  minHoursBefore: 2,       // 何時間前まで変更・取消可能か
  manageUrl: APP_BASE + "reserve.html",
  shopMail:  ""            // 空なら Apps Script 所有者のメールに店側通知
};

// ── ステータス定義 ────────────────────────────────────────────────
const RES_STATUS = {
  CONFIRMED: "✅ 確定 / Potwierdzona",
  UPDATED:   "🔄 変更済 / Zmieniona",
  CANCELLED: "❌ キャンセル / Anulowana",
  ARRIVED:   "🍜 来店済 / Zrealizowana",
  NOSHOW:    "👻 No-show"
};

// ════════════════════════════════════════════════════════════════════
//  セットアップ
// ════════════════════════════════════════════════════════════════════

function setupReservationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (sh) { Logger.log("既に存在します"); return sh; }
  sh = ss.insertSheet(SHEET_RESERVATIONS);
  const headers = [
    "予約ID / ID", "ステータス / Status", "予約日 / Data", "時間 / Godzina",
    "人数 / Osoby", "氏名 / Imię i nazwisko", "Email", "電話 / Telefon",
    "コメント / Uwagi", "会員ID / Member", "ランク / Rank", "プロモ同意 / Promo",
    "トークン(内部用)", "受付日時 / Utworzono", "更新日時 / Zmieniono", "スタッフメモ / Notatka"
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground("#c0392b").setFontColor("#ffffff").setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 170); sh.setColumnWidth(6, 160);
  sh.setColumnWidth(9, 220); sh.setColumnWidth(16, 200);
  sh.hideColumns(13); // トークン列は隠す
  Logger.log("予約シートを作成しました");
  return sh;
}

function setupReservationFormTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 二重登録防止
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onReservationFormSubmit") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onReservationFormSubmit")
    .forSpreadsheet(ss).onFormSubmit().create();
  Logger.log("フォーム送信トリガーを設置しました");
}

// ════════════════════════════════════════════════════════════════════
//  フォーム送信 → 予約登録
// ════════════════════════════════════════════════════════════════════

function onReservationFormSubmit(e) {
  try {
    const nv = e.namedValues || {};
    const entries = Object.keys(nv).map(k => ({ key: k, val: (nv[k] || []).join(", ").trim() }));
    const isConsent = k => /zgod|rodo|consent|agree|akcept|polityk|privacy|regulamin|warunk|terms/i.test(k);
    const isTimestamp = k => /sygnatura|timestamp/i.test(k);

    // 質問タイトル(keyRe)と回答の形式(valRe)の両方で照合。
    // 同意系の質問（RODO・規約など）は除外し、タイトルで見つからない場合は
    // 全回答から形式一致する値を探すフォールバック付き。
    const pick = (keyRe, valRe) => {
      for (const en of entries) {
        if (!en.val || isConsent(en.key) || isTimestamp(en.key)) continue;
        if (keyRe.test(en.key) && (!valRe || valRe.test(en.val))) return en.val;
      }
      if (valRe) {
        for (const en of entries) {
          if (!en.val || isConsent(en.key) || isTimestamp(en.key)) continue;
          if (valRe.test(en.val)) return en.val;
        }
      }
      return "";
    };
    // 同意系質問専用（プロモ同意はこちらで拾う）
    const pickConsent = (keyRe) => {
      for (const en of entries) {
        if (en.val && keyRe.test(en.key)) return en.val;
      }
      return "";
    };

    const email   = pick(/e-?mail|adres/i, /@/);
    const phone   = pick(/telefon|phone/i, /\d[\d\s()+-]{5,}/);
    const dateRaw = pick(/data|date/i, /^\s*(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[.\/]\d{1,2}[.\/]\d{2,4})\s*$/);
    let   time    = normalizeTime(pick(/godzin|time|hour/i, /^\s*\d{1,2}[:.]\d{2}\s*$/));
    const guests  = (pick(/osób|osob|guest|liczba|person/i, /^\s*\d{1,2}\b/).match(/\d+/) || [""])[0];
    const name    = pick(/imię|nazwisko|name/i);
    const comment = pick(/uwag|koment|comment|wiadom|message/i);
    const promo   = /tak|yes|zgadzam|wyrażam/i.test(pickConsent(/promoc|market|newsletter|ofert/i)) ? "○" : "";

    // 最終防壁：時刻形式でなければ空にする（同意文などの誤記入をシートに残さない）
    if (!/^\d{1,2}:\d{2}$/.test(time)) {
      Logger.log("⚠️ 時間を特定できませんでした。回答: " + JSON.stringify(nv));
      time = "";
    }

    if (!email || !dateRaw) { Logger.log("必須項目が特定できません: " + JSON.stringify(nv)); return; }

    // v6.9：登録処理は createReservation() に集約（HP直送信と共通）
    createReservation({
      email, name, phone,
      date: normalizeDate(dateRaw),
      time, guests, comment, promo,
      source: "form"
    });
  } catch (err) {
    Logger.log("onReservationFormSubmit error: " + err);
  }
}

// ── ID・トークン生成 ─────────────────────────────────────────────
function makeReservationId(dateStr) {
  const d = dateStr.replace(/-/g, "");
  const rand = Math.floor(100 + Math.random() * 900);
  return "RSV-" + d + "-" + rand;
}
function makeToken() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let t = "";
  for (let i = 0; i < 8; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

// ── 日付/時刻の正規化 ─────────────────────────────────────────────
function normalizeDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Europe/Warsaw", "yyyy-MM-dd");
  const s = v.toString().trim();
  let m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);               // 2026-07-15
  if (m) return m[1] + "-" + pad2(m[2]) + "-" + pad2(m[3]);
  m = s.match(/(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);           // 15.07.2026 / 15/07/2026
  if (m) return m[3] + "-" + pad2(m[2]) + "-" + pad2(m[1]);
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, "Europe/Warsaw", "yyyy-MM-dd");
  return s;
}
function normalizeTime(v) {
  const m = v.toString().match(/(\d{1,2})[:.](\d{2})/);
  return m ? pad2(m[1]) + ":" + m[2] : v.toString().trim();
}
function pad2(n) { return ("0" + n).slice(-2); }

// ── 予約行の検索 ─────────────────────────────────────────────────
function findReservationRow(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim() === id.trim()) {
      return { sh, row: i + 1, data: data[i] };
    }
  }
  return null;
}

function rowToObj(r) {
  return {
    id: r[0], status: r[1],
    date: r[2] instanceof Date ? Utilities.formatDate(r[2], "Europe/Warsaw", "yyyy-MM-dd") : (r[2] || "").toString(),
    time: r[3] instanceof Date ? Utilities.formatDate(r[3], "Europe/Warsaw", "HH:mm") : (r[3] || "").toString(),
    guests: (r[4] || "").toString(), name: r[5] || "", email: r[6] || "", phone: (r[7] || "").toString(),
    comment: r[8] || "", memberId: r[9] || "", memberRank: r[10] || "",
    createdAt: (r[13] || "").toString(), updatedAt: (r[14] || "").toString(), memo: r[15] || ""
  };
}

// ════════════════════════════════════════════════════════════════════
//  客用API
// ════════════════════════════════════════════════════════════════════

// GET action=getReservation&id=RSV-...&t=TOKEN（またはemail=... 照合）
function getReservationPublic(id, token, email) {
  if (!id) return { status: "error", message: "IDが必要です / Reservation ID required" };
  const found = findReservationRow(id);
  if (!found) return { status: "notFound" };
  const tokenOK = token && (found.data[12] || "").toString().trim() === token.trim();
  const emailOK = email && (found.data[6] || "").toString().trim().toLowerCase() === email.trim().toLowerCase();
  if (!tokenOK && !emailOK) return { status: "unauthorized" };
  const obj = rowToObj(found.data);
  delete obj.memo; // スタッフメモは客に出さない
  obj.canModify = canModifyReservation(obj);
  obj.status = obj.status.toString();
  obj.token = (found.data[12] || "").toString();
  return { status: "ok", reservation: obj, config: RES_CONFIG };
}

function canModifyReservation(obj) {
  if (obj.status.indexOf("キャンセル") >= 0 || obj.status.indexOf("来店済") >= 0 || obj.status.indexOf("No-show") >= 0) return false;
  const dt = new Date(obj.date + "T" + (obj.time || "13:00") + ":00");
  if (isNaN(dt)) return true;
  return (dt.getTime() - Date.now()) > RES_CONFIG.minHoursBefore * 3600 * 1000;
}

// POST action=cancelReservation {id, token}
function handleCancelReservation(data) {
  const found = verifyCustomer(data);
  if (found.error) return jsonResponse(found.error);
  const obj = rowToObj(found.data);
  if (!canModifyReservation(obj)) {
    return jsonResponse({ status: "tooLate", message: "直前のため電話でご連絡ください / Prosimy o kontakt telefoniczny" });
  }
  found.sh.getRange(found.row, 2).setValue(RES_STATUS.CANCELLED);
  found.sh.getRange(found.row, 15).setValue(nowStamp());
  obj.status = RES_STATUS.CANCELLED;
  obj.token = data.token;
  sendReservationMail("cancel", obj);
  notifyShop("予約キャンセル / Anulowanie", obj);
  clearReservationsCache();
  updateTodaySheetSafe();
  return jsonResponse({ status: "ok", message: "キャンセル完了 / Rezerwacja anulowana" });
}

// POST action=updateReservation {id, token, date, time, guests, comment}
function handleUpdateReservation(data) {
  const found = verifyCustomer(data);
  if (found.error) return jsonResponse(found.error);
  const before = rowToObj(found.data);
  if (!canModifyReservation(before)) {
    return jsonResponse({ status: "tooLate", message: "直前のため電話でご連絡ください / Prosimy o kontakt telefoniczny" });
  }

  const newDate   = normalizeDate(data.date || before.date);
  const newTime   = normalizeTime(data.time || before.time);
  const newGuests = (data.guests || before.guests).toString();
  const newComment = data.comment !== undefined ? data.comment : before.comment;

  const check = validateSlot(newDate, newTime, newGuests);
  if (check) return jsonResponse({ status: "invalid", message: check });

  found.sh.getRange(found.row, 2).setValue(RES_STATUS.UPDATED);
  found.sh.getRange(found.row, 3).setValue(newDate);
  found.sh.getRange(found.row, 4).setValue(newTime);
  found.sh.getRange(found.row, 5).setValue(newGuests);
  found.sh.getRange(found.row, 9).setValue(newComment);
  found.sh.getRange(found.row, 15).setValue(nowStamp());

  const after = rowToObj(found.data);
  after.status = RES_STATUS.UPDATED;
  after.date = newDate; after.time = newTime; after.guests = newGuests; after.comment = newComment;
  after.token = data.token;
  sendReservationMail("update", after, before);
  notifyShop("予約変更 / Zmiana rezerwacji", after);
  clearReservationsCache();
  updateTodaySheetSafe();
  return jsonResponse({ status: "ok", message: "変更完了 / Rezerwacja zmieniona", reservation: after });
}

function verifyCustomer(data) {
  if (!data.id || !data.token) return { error: { status: "error", message: "id/token required" } };
  const found = findReservationRow(data.id);
  if (!found) return { error: { status: "notFound" } };
  if ((found.data[12] || "").toString().trim() !== data.token.trim()) {
    return { error: { status: "unauthorized" } };
  }
  return found;
}

// ── 予約枠の妥当性チェック ────────────────────────────────────────
function validateSlot(dateStr, timeStr, guests) {
  const dt = new Date(dateStr + "T" + timeStr + ":00");
  if (isNaN(dt)) return "日付・時間の形式が正しくありません / Nieprawidłowa data";
  if (dt.getTime() < Date.now()) return "過去の日時は選べません / Data w przeszłości";
  const day = dt.getDay();
  if (RES_CONFIG.closedDays.indexOf(day) >= 0) return "定休日です / Zamknięte w tym dniu";
  if (timeStr < RES_CONFIG.openTime || timeStr > RES_CONFIG.lastTime) {
    return "営業時間外です（" + RES_CONFIG.openTime + "–" + RES_CONFIG.lastTime + "）/ Poza godzinami";
  }
  if (timeStr >= RES_CONFIG.lateFrom && RES_CONFIG.lateDays.indexOf(day) < 0) {
    return RES_CONFIG.lateFrom + "以降は金・土のみ / Po " + RES_CONFIG.lateFrom + " tylko pt–sob";
  }
  const g = parseInt(guests, 10);
  if (!g || g < 1) return "人数を選択してください / Wybierz liczbę osób";
  if (g > RES_CONFIG.maxGuests) return "8名以上はお電話でどうぞ / Grupy 8+ tylko telefonicznie";
  return null;
}

// ════════════════════════════════════════════════════════════════════
//  スタッフ用API
// ════════════════════════════════════════════════════════════════════

// GET action=getReservations&password=...&range=today|upcoming|all
//
// v5.0：①30秒キャッシュ（範囲ごと。予約の追加・変更・操作時は即時クリア）
//        ②顧客履歴の集計を各予約に付与：
//          hVisits=通算来店済回数 / hNoshows=通算No-show回数 /
//          hFirst=初回予約か / hLastVisit=最終来店日
//        （メール一致で同一客と判定。メールが空なら電話番号で判定）
function getReservationsForStaff(password, range) {
  const auth5 = verifyStaffPassword(password);
  if (auth5 !== "ok") return { status: "unauthorized", message: STAFF_AUTH_MSG[auth5] };

  const cacheKey = "resv_" + (range || "today");
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (!sh) return { status: "ok", reservations: [] };
  const data = sh.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");

  // 顧客キー：メール（小文字）優先、なければ電話番号
  // 電話は末尾9桁で照合（+48や0048の有無・スペースを吸収。PLの番号は9桁）
  const custKey = o => {
    const e = (o.email || "").toString().trim().toLowerCase();
    if (e) return "e:" + e;
    const p = (o.phone || "").toString().replace(/\D/g, "").slice(-9);
    return p ? "p:" + p : "";
  };

  // ── 1周目：全行から顧客ごとの履歴を集計（どうせ全行読むのでコスト増ほぼゼロ）──
  const hist = {};
  const all = [];
  for (let i = 1; i < data.length; i++) {
    const o = rowToObj(data[i]);
    if (!o.id) continue;
    all.push(o);
    const k = custKey(o);
    if (!k) continue;
    const h = hist[k] || (hist[k] = { visits: 0, noshows: 0, total: 0, lastVisit: "" });
    h.total++;
    if (o.status.toString().indexOf("来店済") >= 0) {
      h.visits++;
      if (o.date > h.lastVisit) h.lastVisit = o.date;
    } else if (o.status.toString().indexOf("No-show") >= 0) {
      h.noshows++;
    }
  }

  // ── 2周目：範囲でフィルタし、履歴を付与 ──
  const list = [];
  all.forEach(o => {
    if (range === "today"    && o.date !== today) return;
    if (range === "upcoming" && o.date <  today) return;
    const h = hist[custKey(o)];
    if (h) {
      o.hVisits    = h.visits;
      o.hNoshows   = h.noshows;
      o.hFirst     = (h.total <= 1);   // 台帳上この1件しかない＝初めてのお客様
      o.hLastVisit = h.lastVisit;
    }
    list.push(o);
  });
  if (range === "all") {
    // 全件：日付の新しい順（同じ日の中は時刻の早い順）
    list.sort((a, b) => b.date.localeCompare(a.date) || (a.time + "").localeCompare(b.time + ""));
  } else {
    // 今日・今後：時系列順
    list.sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
  }

  const result = { status: "ok", reservations: list, today };
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 30); // 30秒
  } catch (e) {} // 100KB超過などで失敗しても動作には影響なし
  return result;
}

// 予約データが変わった時にキャッシュを即時破棄する
// （これを呼ばないと「操作したのに画面が30秒古いまま」になる）
function clearReservationsCache() {
  try {
    CacheService.getScriptCache().removeAll(["resv_today", "resv_upcoming", "resv_all"]);
  } catch (e) {}
}

// POST action=staffReservation {password, id, op, memo}
//   op: confirm | arrived | noshow | cancel
function handleStaffReservation(data) {
  const auth6 = verifyStaffPassword(data.password);
  if (auth6 !== "ok") return jsonResponse({ status: "unauthorized", message: STAFF_AUTH_MSG[auth6] });
  const found = findReservationRow(data.id || "");
  if (!found) return jsonResponse({ status: "notFound" });

  const map = { confirm: RES_STATUS.CONFIRMED, arrived: RES_STATUS.ARRIVED,
                noshow: RES_STATUS.NOSHOW, cancel: RES_STATUS.CANCELLED };
  if (data.op && map[data.op]) {
    found.sh.getRange(found.row, 2).setValue(map[data.op]);
    found.sh.getRange(found.row, 15).setValue(nowStamp());
    if (data.op === "cancel") {
      const obj = rowToObj(found.data);
      obj.status = RES_STATUS.CANCELLED;
      obj.token = (found.data[12] || "").toString();
      sendReservationMail("cancel", obj);
    }
  }
  if (data.memo !== undefined) found.sh.getRange(found.row, 16).setValue(data.memo);
  clearReservationsCache();
  updateTodaySheetSafe();
  return jsonResponse({ status: "ok" });
}

function nowStamp() {
  return Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd HH:mm");
}

// ════════════════════════════════════════════════════════════════════
//  メール
// ════════════════════════════════════════════════════════════════════

function reservationSummary(r) {
  return [
    "🗓 " + r.date + "  🕐 " + r.time,
    "👥 " + r.guests + " os. /名",
    "👤 " + r.name,
    "📞 " + r.phone,
    r.comment ? "📝 " + r.comment : "",
    "🔖 " + r.id
  ].filter(Boolean).join("\n");
}

function sendReservationMail(type, r, before) {
  if (!r.email) return;
  const manage = RES_CONFIG.manageUrl + "?id=" + encodeURIComponent(r.id) + "&t=" + encodeURIComponent(r.token || "");
  let subject, headline;
  if (type === "confirm") {
    subject = "🍜 Wabi Navi — Potwierdzenie rezerwacji / Reservation Confirmed / ご予約確定";
    headline = "Dziękujemy za rezerwację!\nYour reservation is confirmed.\nご予約ありがとうございます。";
  } else if (type === "update") {
    subject = "🔄 Wabi Navi — Rezerwacja zmieniona / Reservation Updated / ご予約変更完了";
    headline = "Twoja rezerwacja została zmieniona.\nYour reservation has been updated.\nご予約内容を変更しました。";
  } else {
    subject = "❌ Wabi Navi — Rezerwacja anulowana / Reservation Cancelled / ご予約キャンセル";
    headline = "Twoja rezerwacja została anulowana.\nYour reservation has been cancelled.\nご予約をキャンセルしました。";
  }

  let body = `
${r.name ? r.name + " 様 / Drogi/a " + r.name + "," : "Drogi Gościu,"}

━━━━━━━━━━━━━━━━━━━━━━━━━━
${headline}
━━━━━━━━━━━━━━━━━━━━━━━━━━

${reservationSummary(r)}
`;
  if (before && type === "update") {
    body += `
(poprzednio / previously / 変更前: ${before.date} ${before.time}, ${before.guests} os.)
`;
  }
  if (type !== "cancel") {
    body += `
🔄 Zmień lub anuluj rezerwację / Change or cancel / 変更・キャンセルはこちら:
${manage}

⏰ Zmiany możliwe do ${RES_CONFIG.minHoursBefore}h przed wizytą.
   Changes possible up to ${RES_CONFIG.minHoursBefore}h before your visit.
   ご来店${RES_CONFIG.minHoursBefore}時間前まで変更可能です。
`;
  } else {
    body += `
Zapraszamy ponownie! / We hope to see you again! / またのご予約をお待ちしております。
`;
  }
  // ── v6.9：会員なら感謝＋おすすめ＋イベント／新規なら登録誘導 ──
  if (type === "confirm") {
    try { body += buildLoyaltyBlock(r); }
    catch (e) { Logger.log("buildLoyaltyBlock: " + e); }
  }
  body += `
📍 Wabi Navi — Toruń
🌐 ${SITE_URL}

━━━━━━━━━━━━━━━━━━━━━━━━━━
和美なび WABI NAVI
━━━━━━━━━━━━━━━━━━━━━━━━━━`;

  try { GmailApp.sendEmail(r.email, subject, body.trim()); }
  catch (e) { Logger.log("reservation mail error: " + e); }
}

function notifyShop(title, r) {
  try {
    const to = RES_CONFIG.shopMail || Session.getEffectiveUser().getEmail();
    if (!to) return;
    GmailApp.sendEmail(to, "🔔 " + title + " — " + r.date + " " + r.time,
      reservationSummary(r) + "\n\n📧 " + r.email +
      (r.memberId ? "\n⭐ 会員: " + r.memberId + " " + (r.memberRank || "") : "\n（非会員 / Gość）"));
  } catch (e) { Logger.log("notifyShop error: " + e); }
}

// ════════════════════════════════════════════════════════════════════
//  本日の予約一覧シート / Dziś — 開店前チェック用
//
//  【セットアップ】setupTodaySheetTrigger() を1回実行
//   → 毎朝6時台に自動更新。さらに予約の新規・変更・取消でも即時更新。
//   → 手動更新はシートのメニュー「🍜 Wabi Navi → 📅 本日の予約を更新」
// ════════════════════════════════════════════════════════════════════

const SHEET_TODAY = "Rezerwacja / Dziś";  // 旧「本日の予約 / Dziś」から改名

function setupTodaySheetTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "updateTodaySheet") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("updateTodaySheet").timeBased().everyDays(1).atHour(6).create();  // 朝
  ScriptApp.newTrigger("updateTodaySheet").timeBased().everyDays(1).atHour(12).create(); // 正午（開店前の最終確認用）
  updateTodaySheet();
  Logger.log("本日予約シートのトリガーを設置しました（毎朝6時台＋正午12時台）");
}

function updateTodaySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_TODAY);
  if (!sh) {
    // 旧名タブがあればリネームして引き継ぐ（データ・位置を保持）
    const legacy = ss.getSheetByName("本日の予約 / Dziś");
    if (legacy) { legacy.setName(SHEET_TODAY); sh = legacy; }
  }
  if (!sh) sh = ss.insertSheet(SHEET_TODAY, 0); // 先頭タブに配置
  sh.clear();
  try { sh.getBandings().forEach(b => b.remove()); } catch (e) {}

  const now = new Date();
  const today = Utilities.formatDate(now, "Europe/Warsaw", "yyyy-MM-dd");
  const dayJp = ["日","月","火","水","木","金","土"][now.getDay()];
  const dayPl = ["niedziela","poniedziałek","wtorek","środa","czwartek","piątek","sobota"][now.getDay()];

  // 本日分を収集
  const resSh = ss.getSheetByName(SHEET_RESERVATIONS);
  const all = [];
  if (resSh) {
    const data = resSh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const o = rowToObj(data[i]);
      if (o.id && o.date === today) all.push(o);
    }
  }
  const isOff = o => /キャンセル|No-show/.test(o.status);
  const byTime = (a, b) => (a.time || "").localeCompare(b.time || "");
  const active = all.filter(o => !isOff(o)).sort(byTime);
  const off    = all.filter(isOff).sort(byTime);
  const totalGuests = active.reduce((s, o) => s + (parseInt(o.guests, 10) || 0), 0);

  const COLS = 8;

  // ── タイトル・サマリー ──
  sh.getRange(1, 1, 1, COLS).merge()
    .setValue("🍜 本日の予約 / Rezerwacje dziś — " + today + "（" + dayJp + " / " + dayPl + "）")
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#c0392b").setFontColor("#ffffff")
    .setVerticalAlignment("middle");
  sh.setRowHeight(1, 40);

  sh.getRange(2, 1, 1, COLS).merge()
    .setValue(active.length === 0
      ? "本日の予約はありません / Brak rezerwacji na dziś 🍃"
      : "📋 予約 " + active.length + " 件　　👥 来店予定 合計 " + totalGuests + " 名")
    .setFontSize(13).setFontWeight("bold")
    .setBackground("#f7e8d8").setFontColor("#2c1810")
    .setVerticalAlignment("middle");
  sh.setRowHeight(2, 32);

  // ── ヘッダー（見てほしい順：時間→人数→氏名→電話） ──
  const headers = ["🕐 Godzina / Time", "👥 Osoby / Guests", "👤 Imię i nazwisko / Name",
                   "📞 Telefon / Phone", "📧 E-mail",
                   "📝 Uwagi / Comments", "⭐ Członek / Member", "🗒 Notatka / Memo"];
  sh.getRange(4, 1, 1, COLS).setValues([headers])
    .setFontWeight("bold").setBackground("#8e2418").setFontColor("#ffffff");
  sh.setFrozenRows(4);

  let row = 5;
  const safeTime = t => ((t || "").toString().match(/\d{1,2}[:.]\d{2}/) || ["⚠️ 要確認"])[0];
  if (active.length) {
    const rows = active.map(o => [
      safeTime(o.time), Number(o.guests) || o.guests, o.name, o.phone, o.email,
      o.comment, o.memberId ? (o.memberRank + " " + o.memberId) : "", o.memo
    ]);
    const rng = sh.getRange(row, 1, rows.length, COLS);
    rng.setNumberFormat("@"); // 電話の先頭0や時刻表記を守る
    rng.setValues(rows).setVerticalAlignment("middle");
    // 主要4列（時間・人数・氏名・電話）を大きく
    sh.getRange(row, 1, rows.length, 4).setFontSize(14).setFontWeight("bold");
    sh.getRange(row, 1, rows.length, 2).setHorizontalAlignment("center");
    sh.getRange(row, 5, rows.length, 4).setFontSize(10).setFontColor("#5a4a3a");
    // 縞模様で読みやすく
    for (let i = 0; i < rows.length; i++) {
      if (i % 2 === 1) sh.getRange(row + i, 1, 1, COLS).setBackground("#fdf3ea");
      sh.setRowHeight(row + i, 30);
    }
    row += rows.length;
  }

  // ── キャンセル・No-show（下部にグレー表示） ──
  if (off.length) {
    row += 1;
    sh.getRange(row, 1, 1, COLS).merge()
      .setValue("― キャンセル・No-show（" + off.length + "件・参考） ―")
      .setFontColor("#999999").setFontSize(10).setHorizontalAlignment("center");
    row += 1;
    const offRows = off.map(o => [
      safeTime(o.time), o.guests, o.name, o.phone, o.email, o.status, "", o.memo
    ]);
    sh.getRange(row, 1, offRows.length, COLS).setNumberFormat("@")
      .setValues(offRows).setFontColor("#aaaaaa").setFontSize(10)
      .setFontLine("line-through");
    row += offRows.length;
  }

  // 列幅（本日一覧とカレンダーの共用）
  const widths = [110, 100, 180, 140, 200, 260, 160, 200];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ── 📆 3週間カレンダー（同一タブ下段・最低21行目から、予約が多い日は自動で繰り下げ） ──
  const calStart = Math.max(21, row + 2);
  renderCalendarSection(sh, calStart, now, today);

  // 旧カレンダータブが残っていれば削除（統合済みのため）
  try {
    const oldCal = ss.getSheetByName("予約カレンダー / Kalendarz");
    if (oldCal) ss.deleteSheet(oldCal);
  } catch (e) {}
}

// カレンダー部の描画（Rezerwacja/Dziśタブの下段）
function renderCalendarSection(sh, startRow, now, today) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 今週の月曜日を起点に3週間
  const monday = new Date(now);
  monday.setDate(now.getDate() - ((now.getDay() + 6) % 7));

  // 有効な予約を日付ごとに収集（キャンセル・No-showは除外）
  const byDate = {};
  const resSh = ss.getSheetByName(SHEET_RESERVATIONS);
  if (resSh) {
    const data = resSh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const o = rowToObj(data[i]);
      if (!o.id || /キャンセル|No-show/.test(o.status)) continue;
      (byDate[o.date] = byDate[o.date] || []).push(o);
    }
  }

  let r = startRow;
  sh.getRange(r, 1, 1, 7).merge()
    .setValue("📆 Kalendarz rezerwacji — 3 tygodnie / 予約カレンダー（3週間）")
    .setFontSize(13).setFontWeight("bold")
    .setBackground("#c0392b").setFontColor("#ffffff").setVerticalAlignment("middle");
  sh.setRowHeight(r, 30);
  r++;

  const dayNames = ["Pon 月", "Wt 火", "Śr 水", "Czw 木", "Pt 金", "Sob 土", "Ndz 日"];
  sh.getRange(r, 1, 1, 7).setValues([dayNames])
    .setFontWeight("bold").setBackground("#8e2418").setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  r++;

  for (let w = 0; w < 3; w++) {
    const values = [], bgs = [];
    for (let d = 0; d < 7; d++) {
      const dt = new Date(monday);
      dt.setDate(monday.getDate() + w * 7 + d);
      const ds = Utilities.formatDate(dt, "Europe/Warsaw", "yyyy-MM-dd");
      const label = Utilities.formatDate(dt, "Europe/Warsaw", "dd.MM");
      const list = (byDate[ds] || []).sort((a, b) => (a.time || "").localeCompare(b.time || ""));
      let txt = label + (ds === today ? " ⭐DZIŚ" : "");
      if (list.length) {
        const g = list.reduce((sum, o) => sum + (parseInt(o.guests, 10) || 0), 0);
        txt += "\n📋 " + list.length + " rez. / " + g + " os.";
        txt += "\n" + list.map(o => "🕐" + (o.time || "?") + " ×" + o.guests).join("\n");
        bgs.push(ds === today ? "#fbd2bd" : "#fdeee6");
      } else {
        txt += "\n—";
        bgs.push(ds === today ? "#fdf3ea" : "#fffcf5");
      }
      values.push(txt);
    }
    sh.getRange(r, 1, 1, 7).setValues([values]).setBackgrounds([bgs])
      .setWrap(true).setVerticalAlignment("top").setFontSize(9);
    sh.setRowHeight(r, 115);
    r++;
  }

  // フッター（タブ全体の最終更新）
  sh.getRange(r, 1, 1, 7).merge()
    .setValue("最終更新 / Aktualizacja: " +
      Utilities.formatDate(now, "Europe/Warsaw", "yyyy-MM-dd HH:mm") +
      " ｜ 手動更新: メニュー「🍜 Wabi Navi → 📅 本日の予約を更新」")
    .setFontSize(9).setFontColor("#7a6858");
}

// 予約の作成・変更・取消時に即時反映（失敗しても本処理は止めない）
function updateTodaySheetSafe() {
  try { updateTodaySheet(); } catch (e) { Logger.log("updateTodaySheet: " + e); }
}

// ════════════════════════════════════════════════════════════════════
//  メニュー更新 2026-07-10 — 丼物・味噌汁セットのトッピング設定
//  【実行方法】この関数を1回実行するとメニューシートに反映されます。
//  ・丼物(o3〜o9)：取消・変更できる具材を たまねぎ/しょうが/ねぎ の3種に統一
//  ・Zestaw Domowy(o11)・ごはんと味噌汁セット(a17)：
//    ねぎを固定具材に追加、わかめ・ゴマを取消可能に、アレルゲンに11(ごま)追加
// ════════════════════════════════════════════════════════════════════
function applyMenuUpdates_donburi_miso() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) { Logger.log("⚠️ メニューシートが見つかりません"); return; }

  const DON_IDS = ["o3","o4","o5","o6","o7","o8","o9"];
  const donRemovable  = "Cebula / Onion:1 | Imbir / Ginger:1 | Szczypiorek / Green onion:1";
  const misoRemovable = "Wakame:2 | Sezam / Sesame:1";

  const data = sh.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][1] || "").toString().trim();
    if (DON_IDS.indexOf(id) >= 0) {
      sh.getRange(i + 1, 14).setValue(donRemovable);           // N列: 抜ける具材
      updated++;
    }
    if (id === "o11") {
      sh.getRange(i + 1, 10).setValue("Ryż, bulion miso z wakame, sezamem i szczypiorkiem");
      sh.getRange(i + 1, 11).setValue("Rice, miso broth with wakame, sesame and green onion");
      sh.getRange(i + 1, 12).setValue("6, 11");                 // L列: アレルゲン
      sh.getRange(i + 1, 14).setValue(misoRemovable);
      updated++;
    }
    if (id === "a17") {
      sh.getRange(i + 1, 10).setValue("Ryż, zupa miso z wakame, sezamem i szczypiorkiem");
      sh.getRange(i + 1, 11).setValue("Rice, miso soup with wakame, sesame and green onion");
      sh.getRange(i + 1, 12).setValue("6, 11");
      sh.getRange(i + 1, 14).setValue(misoRemovable);
      updated++;
    }
  }
  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}
  Logger.log("✅ メニュー更新完了: " + updated + " 品（キャッシュもクリア済み）");
}

// ════════════════════════════════════════════════════════════════════
//  📆 予約カレンダー（3週間ビュー） — 開店前・週間の予約把握用
//  「本日の予約」と同じタイミングで自動更新（朝・正午・予約変動時）
//  手動更新：updateReservationCalendar() を実行、または本日シート更新時に自動
// ════════════════════════════════════════════════════════════════════
// （統合済み）カレンダーは Rezerwacja/Dziś タブの下段に描画されます。
// 互換のため旧関数は updateTodaySheet の呼び出しに転送します。
function updateReservationCalendar() { updateTodaySheet(); }
function updateReservationCalendarSafe() {
  try { updateTodaySheet(); } catch (e) { Logger.log("calendar: " + e); }
}

// ════════════════════════════════════════════════════════════════════
//  📷 商品写真の掲載チェック — メニューシートR列に✅/—を記入
//  【実行方法】写真をアップロードした後に checkMenuPhotos() を実行
//  GitHub Pagesの images/商品ID.jpg の存在を確認します
// ════════════════════════════════════════════════════════════════════
const PHOTO_BASE_URL = APP_BASE + "images/";

function checkMenuPhotos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) { Logger.log("⚠️ メニューシートが見つかりません（先に setupMenuSheet を実行）"); return; }

  sh.getRange(1, 18).setValue("写真 / Photo")
    .setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e")
    .setHorizontalAlignment("center").setFontSize(9);

  const data = sh.getDataRange().getValues();
  const ids = [], rows = [];
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][1] || "").toString().trim();
    if (id) { ids.push(id); rows.push(i + 1); }
  }

  // 30件ずつまとめて存在確認（Rangeヘッダーで1バイトだけ取得＝高速）
  const codes = [];
  for (let start = 0; start < ids.length; start += 30) {
    const chunk = ids.slice(start, start + 30).map(id => ({
      url: PHOTO_BASE_URL + id + ".jpg",
      muteHttpExceptions: true,
      headers: { Range: "bytes=0-0" }
    }));
    UrlFetchApp.fetchAll(chunk).forEach(r => codes.push(r.getResponseCode()));
  }

  let have = 0;
  const marks = rows.map((r, k) => {
    const ok = (codes[k] === 200 || codes[k] === 206);
    if (ok) have++;
    return [ok ? "✅" : "—"];
  });
  sh.getRange(rows[0], 18, marks.length, 1).setValues(marks).setHorizontalAlignment("center");
  sh.setColumnWidth(18, 70);
  Logger.log("✅ 写真チェック完了: " + have + " / " + rows.length + " 品に写真あり（R列に記入済み）");
}

// ════════════════════════════════════════════════════════════════════
//  🍙 おにぎり a〜k 11種の選択肢をメニューシートに反映
//  【実行方法】applyOnigiriOptions() を1回実行
//  o1（1個）＝1種選択、o1b（2個）＝2種選択のプルダウンが追加されます
// ════════════════════════════════════════════════════════════════════
function applyOnigiriOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) { Logger.log("⚠️ メニューシートが見つかりません"); return; }

  const KINDS = ["a. Wodorosty / Seaweed / 海苔", "b. Kimchi 🌶️ / キムチ",
    "c. Mięso mielone / Minced pork / そぼろ肉", "d. Surimi-mayo / カニカマ・マヨ",
    "e. Tuńczyk-mayo / Tuna-mayo / ツナマヨ", "f. Krewetki-mayo / Shrimp-mayo / エビマヨ",
    "g. Warzywa / Veggies / 野菜", "h. Shiso czerwone / Red shiso / 赤しそ",
    "i. Bonito / Katsuo / おかか（鰹節）", "j. Czosnek-chili 🌶️ / Garlic-chili / ガーリックチリ",
    "k. Jajko-Nori / Egg-Nori / 玉子・海苔"];
  const ZEROS = KINDS.map(() => 0);

  const o1Options = JSON.stringify([
    { name: "Rodzaj / Kind / 種類 (a–k)", type: "change", options: KINDS, priceDeltas: ZEROS }
  ]);
  const o1bOptions = JSON.stringify([
    { name: "1. Onigiri — rodzaj / 1個目 (a–k)", type: "change", options: KINDS, priceDeltas: ZEROS },
    { name: "2. Onigiri — rodzaj / 2個目 (a–k)", type: "change", options: KINDS, priceDeltas: ZEROS }
  ]);

  const data = sh.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][1] || "").toString().trim();
    if (id === "o1")  { sh.getRange(i + 1, 15).setValue(o1Options);  updated++; }
    if (id === "o1b") { sh.getRange(i + 1, 15).setValue(o1bOptions); updated++; }
  }
  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}
  Logger.log("✅ おにぎり選択肢を反映: " + updated + " 品（キャッシュもクリア済み）");
}

// ════════════════════════════════════════════════════════════════════
//  🍜✌️ ラーメン全品に「2杯に分割」オプションを追加
//  【実行方法】applyRamenSplitOption() を1回実行
//  既存のオプション（替え玉など）の後ろに追加されます（重複追加は自動回避）
// ════════════════════════════════════════════════════════════════════
function applyRamenSplitOption() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) { Logger.log("⚠️ メニューシートが見つかりません"); return; }

  const SPLIT_GROUP = {
    name: "🍜✌️ Jedna porcja na dwie miseczki / One portion, 2 bowls / 一人前を2杯に分割（ファミリー・カップルにお勧め！）",
    type: "add",
    options: ["Nie / No（standard）", "Tak, podziel! / Yes, split! — polecane dla par i rodzin"],
    priceDeltas: [0, 0]
  };

  const data = sh.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const catId = (data[i][0] || "").toString().trim();
    if (catId !== "ramen") continue;
    let tops = [];
    const raw = (data[i][14] || "").toString().trim(); // O列
    if (raw) {
      try { tops = JSON.parse(raw); } catch (e) { Logger.log("O列JSON不正 行" + (i + 1) + ": " + e); continue; }
    }
    const idx = tops.findIndex(t => (t.name || "").indexOf("dwie miseczki") >= 0);
    if (idx >= 0) tops[idx] = SPLIT_GROUP;   // 旧ラベルがあれば上書き
    else tops.push(SPLIT_GROUP);             // なければ追加
    sh.getRange(i + 1, 15).setValue(JSON.stringify(tops));
    updated++;
  }
  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}
  Logger.log("✅ 一人前2杯分割オプション: " + updated + " 品に反映・キャッシュクリア済み");
}

// ════════════════════════════════════════════════════════════════════
//  💳 会計モジュール（v4.6）
//  ・lookupNip: ポーランド財務省の公式VAT登録簿API（Wykaz podatników VAT）
//    で NIP を照会し、社名・住所・VAT状態を返す
//  ・handleCheckout: 客の「レジ会計」リクエストを注文シートに紫色で記録
// ════════════════════════════════════════════════════════════════════

function lookupNip(nip) {
  nip = (nip || "").toString().replace(/[^0-9]/g, "");
  // 形式＋チェックサム検証
  if (nip.length !== 10) return { status: "invalid" };
  const w = [6, 5, 7, 2, 3, 4, 5, 6, 7];
  let sum = 0;
  for (let i = 0; i < 9; i++) sum += w[i] * Number(nip[i]);
  if (sum % 11 !== Number(nip[9])) return { status: "invalid" };

  try {
    const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
    const url = "https://wl-api.mf.gov.pl/api/search/nip/" + nip + "?date=" + today;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { status: "notFound" };
    const j = JSON.parse(res.getContentText());
    const subj = j && j.result && j.result.subject;
    if (!subj) return { status: "notFound" };
    return {
      status: "ok",
      name: subj.name || "",
      address: subj.workingAddress || subj.residenceAddress || "",
      statusVat: subj.statusVat || ""
    };
  } catch (e) {
    Logger.log("lookupNip: " + e);
    return { status: "error" };
  }
}

function handleCheckout(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) sh = createOrderSheet(ss);

  const invoiceLabel = data.invoiceType === "faktura" ? "📄 FAKTURA VAT"
    : data.invoiceType === "nip" ? "🧾 PARAGON z NIP" : "🧾 Paragon zwykły";
  const nipInfo = data.nip
    ? "NIP: " + data.nip + " ｜ " + (data.companyName || "") + " ｜ " + (data.companyAddress || "")
    : "";
  // v8.1：客ごとの支払額（この卓の全注文の合計）をレジ用に載せる
  const gb = (data.guestBreakdown || "").toString();
  const memo = "💳 PROŚBA O RACHUNEK / お会計リクエスト\n" +
    invoiceLabel + "\n" +
    "Zamówienia / 対象注文: " + ((data.orders || []).join(", ") || "—") + "\n" +
    (gb ? "👤 Do zapłaty / 客ごとの支払額: " + gb + "\n" : "") +
    (nipInfo ? nipInfo + "\n" : "") +
    (data.isMember ? "⭐ Członek: " + (data.memberName || "") + " (" + (data.memberContact || "") + ")" : "Gość / Guest");

  // v8.0：会計リクエストも一番上に差し込む
  const row = prependOrderRows_(sh, [[
    "💳 DO KASY", new Date(), "Stolik / Table " + data.table,
    (data.guests || "") + " os.", "",
    "── PŁATNOŚĆ / お会計 ──", "CHECKOUT REQUEST", "", "", "",
    memo, invoiceLabel + (data.nip ? " (NIP " + data.nip + ")" : ""),
    data.isMember ? "✅ Członek" : "Gość",
    data.memberName || "", data.memberContact || "", "", "",
    "DO KASY / 会計待ち", data.sum || "", ""
  ]]);
  // 紫色で強調（会計リクエストと一目で分かるように）
  sh.getRange(row, 1, 1, 19).setBackground("#e8dff5").setFontWeight("bold");
  sh.getRange(row, 1).setFontColor("#5b3b8e");
  sh.getRange(row, 19).setFontColor("#5b3b8e").setFontSize(12);

  return jsonResponse({ status: "ok" });
}

// ════════════════════════════════════════════════════════════════════
//  🍜 メニュー更新 2026-07-15 — 実行方法: applyMenuUpdates_20260715() を1回実行
//  ・Szczypiorek → Por（ねぎ）に名称統一（全列）
//  ・Dodatkiに Memma（メンマ）・Grzyby Shiitake（しいたけ）を追加（各2zł）
//  ・丼物ページに Zupa Miso（味噌汁 12zł）を追加
//  ・ライス（o10/a16）の説明に「ラーメンとセットで6zł」を追記
// ════════════════════════════════════════════════════════════════════
function applyMenuUpdates_20260715() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) { Logger.log("⚠️ メニューシートが見つかりません"); return; }

  const data = sh.getDataRange().getValues();
  const ids = data.map(r => (r[1] || "").toString().trim());
  let renamed = 0;

  // ① ねぎの名称統一（C,D,H,I,J,K,N列）
  const renameCols = [3, 4, 8, 9, 10, 11, 14];
  for (let i = 1; i < data.length; i++) {
    for (const col of renameCols) {
      let v = (data[i][col - 1] || "").toString();
      if (!v) continue;
      const nv = v
        .replace(/Szczypiorek \/ Green onion/g, "Por / Leek")
        .replace(/Szczypiorek \/ Chives/g, "Por / Leek")
        .replace(/szczypiorkiem/g, "porem")
        .replace(/Szczypiorek/g, "Por")
        .replace(/szczypiorek/g, "por")
        .replace(/[Gg]reen [Oo]nions?/g, "leek")
        .replace(/Chives/g, "Leek").replace(/chives/g, "leek");
      if (nv !== v) { sh.getRange(i + 1, col).setValue(nv); renamed++; }
    }
    // ② ライスにセット案内を追記（未追記の場合のみ）
    const id = ids[i];
    if (id === "o10" || id === "a16") {
      const dPL = (data[i][7] || "").toString();
      if (dPL.indexOf("6 zł") < 0) {
        sh.getRange(i + 1, 8).setValue(dPL + " 🍜 Z ramenem w zestawie: tylko 6 zł (−50%)!");
        sh.getRange(i + 1, 9).setValue((data[i][8] || "") + " 🍜 With any ramen: only 6 zł (half price)!");
      }
    }
  }

  // ③ 新商品の追加（既存チェック付き）
  if (ids.indexOf("d_2g") < 0) {
    sh.appendRow(["dodatki","d_2g","Memma (Menma)","Menma Bamboo","メンマ",2,"表示",
      "Marynowane pędy bambusa.","Seasoned bamboo shoots.","Pędy bambusa","Bamboo shoots","",0,"","","",false]);
  }
  if (ids.indexOf("d_2h") < 0) {
    sh.appendRow(["dodatki","d_2h","Grzyby Shiitake","Shiitake Mushrooms","しいたけ",2,"表示",
      "Grzyby shiitake.","Shiitake mushrooms.","Grzyby shiitake","Shiitake","",0,"","","",false]);
  }
  if (ids.indexOf("o12") < 0) {
    sh.appendRow(["onigiri","o12","Zupa Miso","Miso Soup","味噌汁",12,"表示",
      "Gorąca zupa miso na bulionie drobiowym.","Hot miso soup on chicken broth.",
      "Bulion drobiowy, pasta miso, wakame, sezam","Chicken broth, miso, wakame, sesame",
      "6, 11",0,"Wakame:2 | Sezam / Sesame:1","","",false]);
  }

  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}
  Logger.log("✅ メニュー更新完了（改名 " + renamed + " セル・新商品3点確認・キャッシュクリア済み）");
}

// ════════════════════════════════════════════════════════════════════
//  🍜 メニュー更新 2026-07-16 — 実行方法: applyMenuUpdates_20260716() を1回実行
//  ・Memma / Shiitake の価格を 3zł に修正
//  ・トッピング全品に日本語名を補完（E列）
//  ・3zł帯に新規4品：もやし・カニカマ・ゴマわかめ・豆板醤ペースト
//  ・トッピング内ライス(d_6d)：単品12zł表記＋セット6zł案内（割引はアプリが自動計算）
//  ※再実行しても安全（既存行は更新・なければ追加）
// ════════════════════════════════════════════════════════════════════
function applyMenuUpdates_20260716() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) { Logger.log("⚠️ メニューシートが見つかりません"); return; }

  const JP = {
    "d_noodle1":"替え玉（1玉）","d_noodle2":"半替え玉","d_noodle3":"麺変更（米麺/うどん/そば）",
    "d_1a":"ねぎ","d_1b":"ごま","d_1c":"のり","d_1d":"コーン","d_1e":"紅生姜",
    "d_2a":"チンゲン菜","d_2b":"ほうれん草","d_2c":"わかめ","d_2d":"にんにく","d_2e":"きくらげ","d_2f":"チーズ",
    "d_2g":"メンマ","d_2h":"しいたけ",
    "d_3a":"味玉（半分）","d_3b":"なると","d_3c":"もやし","d_3d":"カニカマ","d_3e":"ゴマわかめ","d_3f":"豆板醤ペースト",
    "d_4a":"キムチ","d_4b":"ひじき枝豆",
    "d_5a":"そぼろ肉","d_5b":"チャーシュー","d_5c":"鶏肉",
    "d_6a":"エビ（3尾）","d_6b":"崩し豆腐","d_6c":"いなり（2個）","d_6d":"ライス（小）","d_6e":"海鮮ミックス"
  };
  const NEW_ROWS = [
    ["dodatki","d_3c","Kiełki Mung","Bean Sprouts","もやし",3,"表示",
      "Chrupiące kiełki fasoli mung.","Crunchy mung bean sprouts.","Kiełki mung","Mung bean sprouts","",0,"","","",false],
    ["dodatki","d_3d","Surimi (Paluszki Krabowe)","Surimi Crab Sticks","カニカマ",3,"表示",
      "Paluszki krabowe surimi.","Surimi crab sticks.","Surimi","Surimi","1, 4",0,"","","",false],
    ["dodatki","d_3e","Goma Wakame","Goma Wakame","ゴマわかめ",3,"表示",
      "Sałatka z wakame z sezamem.","Wakame salad with sesame.","Wakame, sezam","Wakame, sesame","11",0,"","","",false],
    ["dodatki","d_3f","Pasta Chili (Toban Djan)","Chili Bean Paste","豆板醤ペースト",3,"表示",
      "Pikantna pasta chili z bobu — dla miłośników ostrości!","Spicy fermented chili bean paste!","Pasta chili, bób","Chili, broad bean paste","6",3,"","","",false]
  ];

  const data = sh.getDataRange().getValues();
  let updated = 0;
  const found = {};
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][1] || "").toString().trim();
    if (!id) continue;
    found[id] = true;
    if (JP[id] && !(data[i][4] || "").toString().trim()) {
      sh.getRange(i + 1, 5).setValue(JP[id]); updated++;      // E列: 日本語名
    }
    if (id === "d_2g" || id === "d_2h") {
      if (Number(data[i][5]) !== 3) { sh.getRange(i + 1, 6).setValue(3); updated++; }  // 3złに修正
    }
    if (id === "d_6d") {
      if (Number(data[i][5]) !== 12) { sh.getRange(i + 1, 6).setValue(12); updated++; }
      sh.getRange(i + 1, 8).setValue("Mała miseczka ryżu. 🍜 Z ramenem w zestawie: tylko 6 zł (−50%)!");
      sh.getRange(i + 1, 9).setValue("Small bowl of rice. 🍜 With any ramen: only 6 zł (half price)!");
    }
  }
  NEW_ROWS.forEach(r => { if (!found[r[1]]) { sh.appendRow(r); updated++; } });

  try { CacheService.getScriptCache().remove("menu_v1"); } catch (e) {}
  Logger.log("✅ トッピング更新完了: " + updated + " 件（日本語名・価格・新規4品）・キャッシュクリア済み");
}


// ════════════════════════════════════════════════════════════════════
//  🌐 ホームページ連動モジュール / Integracja ze stroną WWW（v6.9）
//
//  目的：
//   ① HPの予約フォームから直接GASへ送信（Googleフォームを経由しない）
//   ② 送信前にメールで会員照合 →「⭐会員を確認しました」を画面表示
//   ③ 確認メールを会員／新規で出し分け
//        会員   → 「毎度、ご愛顧ありがとうございます。」＋おすすめ＋イベント
//        新規客 → 会員登録ページへの誘導＋おすすめ＋イベント（会員限定を除く）
//
//  【初回セットアップ】
//    1. setupEventsSheet() を実行（イベントシート作成）
//    2. ウェブアプリを「デプロイを管理 → 新しいバージョン」で再デプロイ
//    ※ Googleフォーム経由の予約も従来どおり動きます（両方受け付け）
// ════════════════════════════════════════════════════════════════════

const SHEET_EVENTS = "イベント / Wydarzenia";
const REGISTER_URL = APP_BASE + "stamp.html";
const CARD_URL     = APP_BASE + "card.html";

// ── イベントシートの作成 ─────────────────────────────────────────
function setupEventsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let sh = ss.getSheetByName(SHEET_EVENTS);
  if (sh) {
    const r = ui.alert("イベントシートは既に存在します",
      "作り直しますか？（現在の内容は消えます）", ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return sh;
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet(SHEET_EVENTS, 2);
  const headers = [
    "タイトルPL / Tytuł",                  // A
    "タイトルJP / 日本語タイトル",           // B
    "開始日 / Od (yyyy-mm-dd)",            // C
    "終了日 / Do (yyyy-mm-dd)",            // D
    "詳細PL / Opis",                       // E
    "詳細JP / 詳細",                        // F
    "リンク / Link（任意）",                 // G
    "会員限定 / Tylko członkowie（○で限定）", // H
    "表示順 / Kolejność"                    // I
  ];
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold")
    .setBackground("#c0392b").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setFontSize(10);
  [230, 200, 150, 150, 300, 300, 220, 200, 110]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);

  // 記入例（そのまま削除してOK）
  sh.appendRow([
    "Wieczór Sake — degustacja", "日本酒の夕べ",
    "2026-09-05", "2026-09-05",
    "Piątkowa degustacja 4 rodzajów sake z przekąskami. Miejsca ograniczone.",
    "金曜夜、日本酒4種の飲み比べ（おつまみ付き）。席数限定。",
    "", "○", 1
  ]);
  sh.getRange(2, 1, 1, headers.length).setFontColor("#999999").setFontStyle("italic");
  sh.getRange(4, 1).setValue("↑ 記入例です。不要なら行ごと削除してください。");

  ui.alert("✅ イベントシートを作成しました",
    "使い方：\n" +
    "・C/D列の期間内だけメールに掲載されます（期間外は自動で消えます）\n" +
    "・H列に「○」を入れると会員のお客様のメールにだけ載ります\n" +
    "・I列の数字が小さい順に並びます\n\n" +
    "掲載先：HPから予約したお客様への確認メール", ui.ButtonSet.OK);
  return sh;
}

// ── 掲載中のイベントをメール用テキストに ─────────────────────────
function getEventsTextForMail(isMember) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_EVENTS);
    if (!sh) return "";
    const today = Utilities.formatDate(new Date(), "Europe/Warsaw", "yyyy-MM-dd");
    const rows  = sh.getDataRange().getValues();
    const out   = [];

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const titlePL = (r[0] || "").toString().trim();
      const titleJP = (r[1] || "").toString().trim();
      if (!titlePL && !titleJP) continue;
      const from = fmtPromoDate(r[2]);
      const to   = fmtPromoDate(r[3]);
      if (from && today < from) continue;   // 掲載前
      if (to   && today > to)   continue;   // 掲載終了 → 自動で消える
      const memberOnly = (r[7] || "").toString().trim() !== "";
      if (memberOnly && !isMember) continue; // 会員限定は非会員に出さない

      const when = from && to && from !== to ? from + " – " + to : (from || to || "");
      let block = "  ▪ " + [titlePL, titleJP].filter(Boolean).join(" / ");
      if (when) block += "\n     🗓 " + when;
      const descPL = (r[4] || "").toString().trim();
      const descJP = (r[5] || "").toString().trim();
      if (descPL) block += "\n     " + descPL;
      if (descJP) block += "\n     " + descJP;
      const link = (r[6] || "").toString().trim();
      if (link) block += "\n     ▶ " + link;
      if (memberOnly) block += "\n     ⭐ tylko dla członków / 会員限定";
      out.push({ order: Number(r[8]) || 99, text: block });
    }
    if (!out.length) return "";
    out.sort((a, b) => a.order - b.order);
    return "🎌 Wydarzenia / イベントのご案内\n" +
           out.slice(0, 4).map(o => o.text).join("\n\n");
  } catch (e) {
    Logger.log("getEventsTextForMail: " + e);
    return "";
  }
}

// ── 掲載中のおすすめ商品をメール用テキストに ─────────────────────
//    データ元は既存の「おすすめ / Polecane」シート（buildPromoCategory を再利用）
function getPromoTextForMail(limit) {
  try {
    const m = getMenuFromSheet();
    if (!m || !m.menu) return "";
    const cat = m.menu.filter(c => c.id === "osusume")[0];
    if (!cat || !cat.items || !cat.items.length) return "";

    const lines = cat.items.slice(0, limit || 3).map(it => {
      const price = it.origPrice
        ? (it.origPrice + " zł → " + it.price + " zł")
        : (it.price + " zł");
      const jp    = it.nameJP ? "（" + it.nameJP + "）" : "";
      let s = "  • " + it.namePL + jp + " — " + price;
      const badge = [it.promoPL, it.promoJP].filter(Boolean).join(" / ");
      if (badge) s += "\n     " + badge;
      if (it.promoUntil) s += "\n     ⏳ do / まで " + it.promoUntil;
      return s;
    });
    return "⭐ Polecane / 今月のおすすめ\n" + lines.join("\n");
  } catch (e) {
    Logger.log("getPromoTextForMail: " + e);
    return "";
  }
}

// ── 会員／新規で出し分けるメールブロック ─────────────────────────
function buildLoyaltyBlock(r) {
  const isMember = !!r.memberId;
  let out = "";

  if (isMember) {
    out += "\n━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
           "⭐ 毎度、ご愛顧ありがとうございます。\n" +
           "   Dziękujemy za Twoją stałą wierność!\n" +
           "   Thank you for your continued support!\n" +
           "━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
           "🪪 " + r.memberId +
           (r.memberRank   ? "  ・  " + r.memberRank : "") +
           (r.memberPoints ? "  ・  " + r.memberPoints + " pt" : "") + "\n" +
           "🎴 Twoja karta / 会員カード: " + CARD_URL + "\n";
  } else {
    out += "\n━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
           "🎁 Dołącz do Wabi Navi Club / 会員登録のご案内\n" +
           "━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
           "Zbieraj punkty przy każdej wizycie, odbieraj kupony\n" +
           "i wypożyczaj japońskie książki za darmo.\n" +
           "Rejestracja zajmuje minutę i jest bezpłatna.\n\n" +
           "Collect points on every visit, receive coupons\n" +
           "and borrow Japanese books for free. Free, takes a minute.\n\n" +
           "ご来店ごとにポイントが貯まり、クーポンや\n" +
           "日本語書籍の無料貸出をご利用いただけます。登録は1分・無料です。\n\n" +
           "▶ " + REGISTER_URL + "\n";
  }

  const promo = getPromoTextForMail(3);
  if (promo) out += "\n" + promo + "\n";
  const ev = getEventsTextForMail(isMember);
  if (ev) out += "\n" + ev + "\n";
  return out;
}

// ── 会員照合API（GDPR配慮：会員か否かだけを返す）────────────────
//    GET action=checkMember&email=...
//    氏名・ポイント・会員IDは返しません（本人確認ができないため）。
//    詳細は本人しか読めない確認メールでお伝えします。
function checkMemberPublic(email) {
  const e = (email || "").toString().trim();
  if (!e || e.indexOf("@") < 0) return { status: "invalid", isMember: false };
  try {
    const cache = CacheService.getScriptCache();
    const key   = "chkm_" + Utilities.base64EncodeWebSafe(e.toLowerCase()).slice(0, 60);
    const hit   = cache.get(key);
    if (hit !== null) return { status: "ok", isMember: hit === "1" };

    const m = findMember(e);
    const isMember = !!m;
    try { cache.put(key, isMember ? "1" : "0", 300); } catch (err) {}
    return { status: "ok", isMember };
  } catch (err) {
    Logger.log("checkMemberPublic: " + err);
    return { status: "error", isMember: false };
  }
}

// ── 予約の共通登録処理 ───────────────────────────────────────────
//    Googleフォーム経由（onReservationFormSubmit）と
//    HP直送信（handleWebReservation）の両方がここを通ります。
function createReservation(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_RESERVATIONS) || setupReservationsSheet();

  const dateStr = normalizeDate(p.date);
  const time    = /^\d{1,2}:\d{2}$/.test(normalizeTime(p.time || "")) ? normalizeTime(p.time) : "";
  const id      = makeReservationId(dateStr);
  const token   = makeToken();
  const now     = new Date();

  // 会員連動：メール一致で会員情報を取得
  let memberId = "", memberRank = "", memberPoints = 0, memberName = "";
  try {
    const m = findMember(p.email);
    if (m) {
      memberId     = m.memberId || "";
      memberRank   = m.rank     || "";
      memberPoints = m.points   || 0;
      memberName   = m.name     || "";
    }
  } catch (err) { Logger.log("member lookup: " + err); }

  sh.appendRow([
    id, RES_STATUS.CONFIRMED, dateStr, time, p.guests, p.name, p.email, p.phone,
    p.comment, memberId, memberRank, p.promo, token,
    Utilities.formatDate(now, "Europe/Warsaw", "yyyy-MM-dd HH:mm"), "",
    p.source === "web" ? "🌐 HP経由 / ze strony WWW" : ""
  ]);

  const res = {
    id, token, date: dateStr, time,
    guests: p.guests, name: p.name || memberName, email: p.email,
    phone: p.phone, comment: p.comment,
    memberId, memberRank, memberPoints
  };

  sendReservationMail("confirm", res);
  notifyShop("新規予約 / Nowa rezerwacja" + (p.source === "web" ? "（HP）" : ""), res);
  clearReservationsCache();
  updateTodaySheetSafe();
  return res;
}

// ── HPの予約フォームからの直接POST ───────────────────────────────
//    POST { action:"webReservation", email, name, phone, date, time,
//           guests, comment, promo, gdpr, hp }
function handleWebReservation(data) {
  try {
    const email = (data.email || "").toString().trim();
    const name  = (data.name  || "").toString().trim();
    const date  = (data.date  || "").toString().trim();
    const time  = normalizeTime((data.time || "").toString().trim());
    const guests = (data.guests || "").toString().replace(/[^\d]/g, "");

    // ハニーポット（botが埋める隠しフィールド）
    if ((data.hp || "").toString().trim() !== "") {
      return jsonResponse({ status: "ok", id: "-", isMember: false });  // 静かに破棄
    }
    if (!data.gdpr) {
      return jsonResponse({ status: "error",
        message: "Wymagana zgoda RODO. / GDPR consent required. / GDPR同意が必要です。" });
    }
    if (!email || email.indexOf("@") < 0) {
      return jsonResponse({ status: "error",
        message: "Podaj poprawny e-mail. / Valid email required. / メールアドレスをご確認ください。" });
    }
    // v6.9：会員なら氏名・電話は会員台帳から補完（お客様の入力を省略できる）
    let member = null;
    try { member = findMember(email); } catch (e) { Logger.log("member lookup: " + e); }
    const finalName = name || (member ? (member.name || "") : "");
    if (!finalName) {
      return jsonResponse({ status: "error",
        message: "Podaj imię i nazwisko. / Name required. / お名前をご入力ください。" });
    }

    const dateStr = normalizeDate(date);
    const slotErr = validateSlot(dateStr, time, guests);
    if (slotErr) return jsonResponse({ status: "error", message: slotErr });

    // 二重送信ガード（同じメール＋同じ日時を90秒以内に再送信した場合）
    const cache = CacheService.getScriptCache();
    const dupKey = "resdup_" +
      Utilities.base64EncodeWebSafe(email.toLowerCase() + dateStr + time).slice(0, 80);
    if (cache.get(dupKey)) {
      return jsonResponse({ status: "duplicate",
        message: "Ta rezerwacja została już wysłana. / このご予約は送信済みです。" });
    }
    try { cache.put(dupKey, "1", 90); } catch (e) {}

    const res = createReservation({
      email, name: finalName,
      phone:   (data.phone   || "").toString().trim(),
      date:    dateStr,
      time:    time,
      guests:  guests,
      comment: (data.comment || "").toString().trim(),
      promo:   data.promo ? "○" : "",
      source:  "web"
    });

    return jsonResponse({
      status: "ok",
      id: res.id,
      isMember: !!res.memberId,
      message: res.memberId
        ? "毎度、ご愛顧ありがとうございます。 / Dziękujemy za stałą wierność!"
        : "Dziękujemy! / ありがとうございます。"
    });
  } catch (err) {
    Logger.log("handleWebReservation: " + err);
    return jsonResponse({ status: "error", message: err.toString() });
  }
}


// ════════════════════════════════════════════════════════════════════
//  🩹 会員シート Q列（来店履歴）の修復 — v7.0
//
//  【背景】v6.9以前は findMember() が来店履歴を I列（性別）から読んでいたため、
//         updateMember() が Q列に「男性 / Male,2026-07-20」のように
//         性別ラベル混じりの値を書き戻していました。
//         その結果、来店回数（P列）が実態より少なく表示され、
//         会員カードのスタンプグリッドも正しく描画されません。
//
//  【この関数がすること】
//    ・Q列から日付（yyyy-MM-dd）以外の値をすべて除去
//    ・重複した日付を1つにまとめ、古い順に並べ直す
//    ・P列（来店回数）を実際の日付件数に再計算
//    ・O列（最終来店日）を履歴の最終日に合わせる
//    ※ ポイント（K列）・ランク・クーポン（R列）には一切触れません
//
//  【実行方法】メニュー「🍜 Wabi Navi → 🩹 来店履歴を修復」
//              または本関数を直接1回実行。何度実行しても安全です。
//
//  ⚠️ 注意：混入期間中に失われた過去の来店日は復元できません
//     （上書きされて残っていないため）。より正確に戻したい場合は、
//     先に repairMemberHistoryFromLog() を実行してください（下記）。
// ════════════════════════════════════════════════════════════════════
function repairMemberHistoryColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getMemberSheet(ss);
  const ui = SpreadsheetApp.getUi();
  const last = sh.getLastRow();
  if (last < 2) { ui.alert("会員データがありません / Brak danych"); return; }

  const rng  = sh.getRange(2, 1, last - 1, 17);
  const data = rng.getValues();
  const isDate = v => /^\d{4}-\d{2}-\d{2}$/.test(v);

  let touched = 0, removed = 0;
  const outHist = [], outCount = [], outLast = [];

  for (let i = 0; i < data.length; i++) {
    const raw = (data[i][16] || "").toString();
    const parts = raw.split(",").map(s => s.trim()).filter(Boolean);
    const dates = [];
    let dirty = false;

    parts.forEach(p => {
      if (isDate(p)) {
        // ⚠️ 同じ日付が2件あるのは正常（+20pt＝ダブル注文の2杯目）。
        //    来店回数は「日数」ではなく「杯数」なので重複を消さないこと。
        dates.push(p);
      } else {
        dirty = true; removed++;               // 性別ラベル等の混入のみ除去
      }
    });
    dates.sort();

    if (dirty) touched++;
    outHist.push([dates.join(",")]);
    outCount.push([dates.length]);
    outLast.push([dates.length ? dates[dates.length - 1] : ""]);
  }

  sh.getRange(2, 17, outHist.length,  1).setValues(outHist);   // Q列 履歴
  sh.getRange(2, 16, outCount.length, 1).setValues(outCount);  // P列 来店回数
  sh.getRange(2, 15, outLast.length,  1).setValues(outLast);   // O列 最終来店日

  ui.alert("✅ 来店履歴の修復が完了しました",
    "対象会員数 / Członków: " + data.length + "\n" +
    "修正した行 / Poprawiono: " + touched + "\n" +
    "除去した不正値 / Usunięto: " + removed + " 件\n\n" +
    "・P列（来店回数）とO列（最終来店日）も再計算しました\n" +
    "・ポイントとクーポンは変更していません\n\n" +
    "より正確に履歴を復元したい場合は、続けて\n" +
    "「🩹 ポイント履歴から来店日を復元」を実行してください。",
    ui.ButtonSet.OK);
  Logger.log("repairMemberHistoryColumn: rows=" + data.length +
             " touched=" + touched + " removed=" + removed);
}

// ════════════════════════════════════════════════════════════════════
//  🩹 ポイント履歴シートから来店日を復元 — v7.0（任意・推奨）
//
//  「ポイント履歴 / Point Log」には全てのポイント変動が日時つきで
//  残っています。そこから「来店＝+10pt が付いた日」を拾い直し、
//  会員シートQ列の来店履歴を再構築します。
//  （クイズの+5pt・+1pt、失効・取消は来店として数えません）
//
//  【実行方法】repairMemberHistoryColumn() の後に1回実行。
//              何度実行しても安全です（常に履歴から作り直します）。
// ════════════════════════════════════════════════════════════════════
function repairMemberHistoryFromLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const log = ss.getSheetByName(SHEET_POINTLOG);
  if (!log) { ui.alert("ポイント履歴シートがありません / Brak arkusza"); return; }

  const r = ui.alert("🩹 ポイント履歴から来店日を復元",
    "「ポイント履歴」シートの +10pt / +20pt の記録から、\n" +
    "各会員の来店日を再構築して会員シートQ列に書き込みます。\n\n" +
    "・現在のQ列の内容は置き換えられます\n" +
    "・ポイント・ランク・クーポンは変更しません\n\n" +
    "実行してよろしいですか？", ui.ButtonSet.OK_CANCEL);
  if (r !== ui.Button.OK) return;

  // ── ポイント履歴から「来店した日」を会員ごとに集計 ──
  const lg = log.getDataRange().getValues();
  const byContact = {};   // 連絡先(小文字) → {日付: true}
  const isVisit = src => {
    const s = (src || "").toString();
    // クイズ・トリビア由来は来店に数えない
    if (/[Qq]uiz|クイズ|日本クイズ|trivia/.test(s)) return false;
    return true;
  };

  for (let i = 1; i < lg.length; i++) {
    const ts      = (lg[i][0] || "").toString();          // A列 日時
    const contact = (lg[i][2] || "").toString().trim().toLowerCase(); // C列 連絡先
    const change  = (lg[i][4] || "").toString();          // E列 変化量
    const source  = (lg[i][7] || "").toString();          // H列 経路
    if (!contact || !ts) continue;

    const m = change.match(/^\+(\d+)pt$/);
    if (!m) continue;                                     // 減算・失効は対象外
    const amt = Number(m[1]);
    if (amt !== 10 && amt !== 20) continue;               // 来店由来は+10/+20のみ
    if (!isVisit(source)) continue;

    const day = ts.slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) continue;

    // 来店回数は「杯数」で数える仕様のため、+20pt（ダブル）は同日2件を積む
    if (!byContact[contact]) byContact[contact] = [];
    byContact[contact].push(day);
    if (amt === 20) byContact[contact].push(day);
  }

  // ── 会員シートへ反映 ──
  const sh   = getMemberSheet(ss);
  const last = sh.getLastRow();
  if (last < 2) { ui.alert("会員データがありません"); return; }
  const data = sh.getRange(2, 1, last - 1, 17).getValues();

  const outHist = [], outCount = [], outLast = [];
  let restored = 0, totalDays = 0;

  for (let i = 0; i < data.length; i++) {
    const contact = (data[i][1] || "").toString().trim().toLowerCase();
    const list = byContact[contact];
    if (!list) {
      // ログに記録のない会員は現状維持（日付だけに整形して残す）
      const keep = (data[i][16] || "").toString().split(",")
        .map(s => s.trim()).filter(s => /^\d{4}-\d{2}-\d{2}$/.test(s));
      keep.sort();
      outHist.push([keep.join(",")]);
      outCount.push([keep.length]);
      outLast.push([keep.length ? keep[keep.length - 1] : ""]);
      continue;
    }
    const days = list.slice().sort();
    restored++;
    totalDays += days.length;
    outHist.push([days.join(",")]);
    outCount.push([days.length]);
    outLast.push([days.length ? days[days.length - 1] : ""]);
  }

  sh.getRange(2, 17, outHist.length,  1).setValues(outHist);
  sh.getRange(2, 16, outCount.length, 1).setValues(outCount);
  sh.getRange(2, 15, outLast.length,  1).setValues(outLast);

  ui.alert("✅ 来店日の復元が完了しました",
    "復元できた会員 / Odtworzono: " + restored + " 名\n" +
    "復元した来店日 / Wizyt razem: " + totalDays + " 日\n\n" +
    "ポイント履歴に記録がない会員は、既存のQ列から\n" +
    "日付だけを残す形に整形しました。\n" +
    "ポイント・ランク・クーポンは変更していません。",
    ui.ButtonSet.OK);
  Logger.log("repairMemberHistoryFromLog: restored=" + restored + " days=" + totalDays);
}


// ════════════════════════════════════════════════════════════════════
//  📸 メール写真アップロード / Zdjęcia przez e-mail（v7.2）
//
//  スタッフが専用Gmailに「件名＝商品ID・写真添付」で送ると、
//  GASが10分ごとに受信 → 縮小 → GitHubの images/ へ自動コミットし、
//  注文アプリとHPの両方に反映する。
//
//  受付ルール（件名で判定）：
//    ・既存商品ID（例 r4）      → images/r4.jpg（★メニュー未登録IDは拒否＝厳格モード）
//    ・promo_で始まる（promo_x） → images/promo/x.jpg（商品ID照合なし・自由な名前）
//    ・上記以外／実在しないID     → 保存せず「未登録です」と返信
//
//  安全策：
//    ・許可した送信元アドレスのメールだけ処理（PHOTO_CONFIG.allowed）
//    ・幅を上限までリサイズ（Drive変換を利用・外部API不要）
//    ・処理済みメールはゴミ箱へ（受信箱が溜まらない）
//    ・完了／エラーを送信者に返信
//
//  【初回セットアップ】メニュー「🍜 Wabi Navi → 📸 …」から
//    ① setupPhotoUpload（GitHub設定・許可アドレス登録）
//    ② setupPhotoTrigger（10分ごとの自動実行をON）
// ════════════════════════════════════════════════════════════════════

// 設定はスクリプトプロパティに保存（コード上にトークンを書かない）
function photoCfg_() {
  const p = PropertiesService.getScriptProperties();
  return {
    owner:  p.getProperty("GH_OWNER")  || "Riichi-Ramen",
    repo:   p.getProperty("GH_REPO")   || "wabinavi-order",
    branch: p.getProperty("GH_BRANCH") || "main",
    token:  p.getProperty("GH_TOKEN")  || "",
    dir:    p.getProperty("PHOTO_DIR") || "images",
    maxW:   Number(p.getProperty("PHOTO_MAXW") || "800"),
    allowed: (p.getProperty("PHOTO_ALLOWED") || "").split(",")
              .map(s => s.trim().toLowerCase()).filter(Boolean),
    label:  "foto-done"
  };
}

// ── 初回設定：GitHub情報と許可アドレスを対話入力 ───────────────
function setupPhotoUpload() {
  const ui = SpreadsheetApp.getUi();
  const p  = PropertiesService.getScriptProperties();

  const tk = ui.prompt("① GitHub Token / GitHubトークン",
    "GitHubのPersonal Access Token（ghp_… または github_pat_…）を貼り付け。\n" +
    "権限は Contents: Read and write のみでOK（リポジトリ1つに限定推奨）。\n\n" +
    "空欄のままOKを押すと現状維持。",
    ui.ButtonSet.OK_CANCEL);
  if (tk.getSelectedButton() !== ui.Button.OK) return;
  const tkv = tk.getResponseText().trim();
  if (tkv) p.setProperty("GH_TOKEN", tkv);

  const al = ui.prompt("② Dozwolone adresy / 許可する送信元アドレス",
    "写真を送ってよいGmailアドレスをカンマ区切りで入力。\n" +
    "ここに無いアドレスからのメールは無視されます（迷惑画像対策）。\n\n" +
    "例：anna@gmail.com, marek@gmail.com\n\n" +
    "現在：" + (p.getProperty("PHOTO_ALLOWED") || "（未設定）"),
    ui.ButtonSet.OK_CANCEL);
  if (al.getSelectedButton() !== ui.Button.OK) return;
  const alv = al.getResponseText().trim();
  if (alv) p.setProperty("PHOTO_ALLOWED", alv);

  // 任意：リポジトリ情報の変更（通常は既定でOK）
  const rp = ui.prompt("③ Repozytorium / リポジトリ（任意）",
    "owner/repo/branch を変える場合のみ入力（例 Riichi-Ramen/wabinavi-order/main）。\n" +
    "空欄OKで既定（Riichi-Ramen / wabinavi-order / main）。",
    ui.ButtonSet.OK_CANCEL);
  if (rp.getSelectedButton() === ui.Button.OK) {
    const v = rp.getResponseText().trim();
    if (v && v.indexOf("/") > 0) {
      const [o, r, b] = v.split("/");
      if (o) p.setProperty("GH_OWNER", o.trim());
      if (r) p.setProperty("GH_REPO", r.trim());
      if (b) p.setProperty("GH_BRANCH", b.trim());
    }
  }

  const c = photoCfg_();
  ui.alert("✅ 設定を保存しました",
    "リポジトリ / Repo: " + c.owner + "/" + c.repo + " (" + c.branch + ")\n" +
    "保存先 / Folder: " + c.dir + "/\n" +
    "画像処理 / Obraz: 原本のまま保存（送信側で縮小推奨）\n" +
    "許可アドレス / Dozwolone: " + (c.allowed.join(", ") || "（未設定）") + "\n" +
    "トークン / Token: " + (c.token ? "設定済み" : "⚠️ 未設定") + "\n\n" +
    "次に「📸 自動受信をON / Włącz odbiór」を実行してください。",
    ui.ButtonSet.OK);
}

// ── 10分ごとの自動受信トリガーをON ───────────────────────────────
function setupPhotoTrigger() {
  const ui = SpreadsheetApp.getUi();
  const c  = photoCfg_();
  if (!c.token) { ui.alert("⚠️ 先に「📸 写真アップロード設定」でトークンを登録してください"); return; }
  if (!c.allowed.length) { ui.alert("⚠️ 許可する送信元アドレスが未設定です"); return; }

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "processPhotoInbox") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("processPhotoInbox").timeBased().everyMinutes(10).create();
  ui.alert("✅ 自動受信をONにしました",
    "10分ごとに受信箱をチェックします。\n\n" +
    "スタッフへの案内：\n" +
    "・宛先：この設定をしたGmail\n" +
    "・件名：foto:商品ID（例 foto:r4）／プロモは promo:名前（例 promo:lato）／\n" +
    "  Polecane専用写真は polecane:商品ID（例 polecane:r4）\n" +
    "・写真を1枚添付して送信\n\n" +
    "テストするには自分で1通送り、10分待つか\n" +
    "processPhotoInbox を手動実行してください。",
    ui.ButtonSet.OK);
}

function stopPhotoTrigger() {
  const ui = SpreadsheetApp.getUi();
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "processPhotoInbox") { ScriptApp.deleteTrigger(t); n++; }
  });
  ui.alert(n ? "🛑 自動受信をOFFにしました" : "自動受信は元々OFFです");
}

// ── 受信箱を処理（トリガー本体・手動実行も可） ───────────────────
function processPhotoInbox() {
  const c = photoCfg_();
  if (!c.token || !c.allowed.length) { Logger.log("photo: 設定未完了"); return; }

  // 未処理（foto-doneラベル無し・添付あり）の受信メールを検索
  // ★未読のみ処理（既に受信箱にある古いメールや、他用途のメールを拾わない）
  //   写真を送ったら未読で届く→処理後に既読化＆ゴミ箱へ、で二度と拾わない
  const threads = GmailApp.search('is:unread has:attachment -label:' + c.label + ' newer_than:2d', 0, 20);
  let label = GmailApp.getUserLabelByName(c.label) || GmailApp.createLabel(c.label);

  threads.forEach(th => {
    th.getMessages().forEach(msg => {
      try {
        const from = extractEmail_(msg.getFrom());
        // 許可送信元チェック
        if (c.allowed.indexOf(from.toLowerCase()) === -1) {
          Logger.log("photo: 未許可の送信元 " + from + " → 無視");
          return;
        }
        const subject = (msg.getSubject() || "").trim();
        const atts = msg.getAttachments()
          .filter(a => /^image\//.test(a.getContentType()));
        if (!atts.length) return;

        // ★写真依頼は件名を合図（プレフィックス）で始めるルール。業務メールと完全分離。
        //   ・商品写真     → foto:商品ID      （例 foto:r4）        images/商品ID.jpg
        //   ・プロモ写真   → promo:名前        （例 promo:lato）     images/promo/名前.jpg
        //   ・Polecane写真 → polecane:商品ID   （例 polecane:r4）    images/promo/商品ID.jpg ★v7.8
        //   合図で始まらないメールは返信もせず静かに既読化して無視する。
        //   受信箱はメイン(wabinavi@gmail.com)と共用のため、この合図が無いと
        //   返品連絡・請求書・転送メール等を誤処理してしまう。
        //
        //   polecane: は「メニュー掲載中の商品を、Polecane欄でだけ違う写真で見せたい」
        //   時に使う。通常メニューの images/商品ID.jpg は変更しない。厳格モード対象
        //   （実在する商品IDのみ）。promo:と違いSNS投稿パッケージは生成しない。
        const raw = subject.replace(/\s+/g, "").toLowerCase();
        let key, kind;   // kind: "foto" | "promo" | "polecane"
        let mp = raw.match(/^promo[:：](.+)$/);          // promo:名前
        let mpc = raw.match(/^polecane[:：](.+)$/);      // polecane:商品ID
        let mf = raw.match(/^foto[:：](.+)$/);           // foto:商品ID（旧 foto:promo_ も許容）
        if (mp) {
          kind = "promo";
          key = "promo_" + mp[1];                        // 内部形式は従来どおり promo_ に統一
        } else if (mpc) {
          kind = "polecane";
          key = mpc[1];
        } else if (mf) {
          key = mf[1];
          kind = key.indexOf("promo_") === 0 ? "promo" : "foto";  // 旧式 foto:promo_lato も引き続き有効
        } else {
          Logger.log("photo: 合図(foto:/promo:/polecane:)で始まらないため無視 → " + subject);
          markDone_(th, label);   // 既読化＆ゴミ箱（次回拾わない）
          return;
        }

        // 保存先パスと表示名を決める
        let path, human;
        if (kind === "promo") {
          const name = key.replace(/^promo_/, "").replace(/[^a-z0-9_-]/g, "");
          if (!name) { replyPhoto_(msg, from, "err", "プロモ名が空です。件名を promo:名前 の形にしてください（例 promo:lato）。"); markDone_(th, label); return; }
          path  = c.dir + "/promo/" + name + ".jpg";
          human = "promo/" + name;
        } else if (kind === "polecane") {
          // ★厳格モード：メニューに実在する商品IDのみ許可（foto:と同じ照合）
          const item = findMenuItemById_(key);
          if (!item) {
            replyPhoto_(msg, from, "err",
              "「" + key + "」はメニューに未登録の商品IDです。\n" +
              "Polecane用の写真も、まず通常メニューに商品を登録してから送ってください。");
            markDone_(th, label);
            return;
          }
          // ★v7.8修正：promo:（SNS宣伝・自由な名前）と保存先が同じだと、
          //   promo:r4 のような送信でPolecane専用写真を誤って上書きしてしまう。
          //   事故を根絶するため専用フォルダ images/polecane/ に分離する。
          path  = c.dir + "/polecane/" + item.id + ".jpg";
          human = "Polecane " + item.id + "（" + (item.namePL || "") + "）";
        } else {
          // ★厳格モード：メニューに実在する商品IDのみ許可
          const item = findMenuItemById_(key);
          if (!item) {
            replyPhoto_(msg, from, "err",
              "「" + key + "」はメニューに未登録の商品IDです。\n" +
              "先にメニューシートへ商品を追加してから、もう一度お送りください。\n" +
              "（プロモ画像なら promo:名前 、Polecane専用写真なら polecane:商品ID にしてください）");
            markDone_(th, label);
            return;
          }
          path  = c.dir + "/" + item.id + ".jpg";
          human = item.id + "（" + (item.namePL || "") + "）";
        }

        // 最初の画像だけ採用（幅をリサイズしてJPEG化）
        const jpg = resizeToJpeg_(atts[0], c.maxW);
        const res = githubPutFile_(c, path, jpg,
          "photo: update " + path + " (via mail from " + from + ")");

        if (res.ok) {
          replyPhoto_(msg, from, "ok",
            "✅ " + human + " の写真を更新しました。\n" +
            "数分後に注文アプリとホームページに反映されます。\n" +
            "Zdjęcie zaktualizowane — pojawi się za kilka minut.");
          // プロモ写真なら、SNS投稿パッケージも別メールで届ける
          // プロモ写真（SNS宣伝用）だけ、SNS投稿パッケージも別メールで届ける。
          // polecane: はPolecane欄専用の写真なのでSNS投稿文は生成しない。
          if (kind === "promo") {
            try {
              const caption = extractCaption_(msg.getPlainBody());
              sendSocialPackage_(c, from, path, caption, atts[0]);
            } catch (e) { Logger.log("social package: " + e); }
          }
        } else {
          replyPhoto_(msg, from, "err",
            "GitHubへの保存に失敗しました（" + res.code + "）。\n" +
            "時間をおいて再送するか、管理者にご連絡ください。");
        }
        markDone_(th, label);
      } catch (e) {
        Logger.log("photo msg error: " + e);
      }
    });
  });
}

// ── メニューから商品IDで1件検索（表示/売切/非表示を問わず実在確認） ──
function findMenuItemById_(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MENU);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const key = (id || "").toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][1] || "").toString().trim().toLowerCase() === key) {
      return { id: (data[i][1] || "").toString().trim(), namePL: data[i][2] || "" };
    }
  }
  return null;
}

//   仕組み：Google SlidesにフルサイズBlobを配置→縮小→PNG書出し。
//   さらにDrive変換でJPEGにする。失敗時は原本をそのまま返す。
// ── 画像をJPEGにして返す（v7.7：確実さ優先で原本そのまま保存） ──
//   旧版はSlidesに貼ってPNG書き出し→リサイズしていたが、Slidesへの画像挿入が
//   非同期のため、描画前にエクスポートされ「真っ白なJPEG」が保存される不具合が出た。
//   リサイズは諦め、送られた写真をそのままJPEGとして保存する。中身が確実に残る。
//   容量が気になる場合は、送信側で縮小してもらう運用にする（マニュアルに明記）。
function resizeToJpeg_(att, maxW) {
  const blob = att.copyBlob();
  const ct = (blob.getContentType() || "").toLowerCase();
  // 既にJPEGならそのまま。他形式（PNG/HEIC等）はJPEGへ変換だけ試みる。
  if (ct === "image/jpeg" || ct === "image/jpg") return blob;
  try { return blob.getAs("image/jpeg"); }
  catch (e) { Logger.log("jpeg convert fallback: " + e); return blob; }
}

// ── GitHub Contents API でファイルを作成/更新（upsert） ──────────
function githubPutFile_(c, path, blob, message) {
  const base = "https://api.github.com/repos/" + c.owner + "/" + c.repo + "/contents/" + path;
  const headers = {
    Authorization: "token " + c.token,
    Accept: "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28"
  };
  // 既存ファイルのSHAを取得（あれば更新、なければ新規）
  let sha = null;
  try {
    const g = UrlFetchApp.fetch(base + "?ref=" + encodeURIComponent(c.branch),
      { headers, muteHttpExceptions: true });
    if (g.getResponseCode() === 200) sha = JSON.parse(g.getContentText()).sha;
  } catch (e) {}

  const payload = {
    message: message,
    content: Utilities.base64Encode(blob.getBytes()),
    branch: c.branch
  };
  if (sha) payload.sha = sha;

  const res = UrlFetchApp.fetch(base, {
    method: "put", headers,
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  return { ok: code === 200 || code === 201, code };
}

// ── 送信者へ返信 ─────────────────────────────────────────────────
function replyPhoto_(msg, to, kind, body) {
  try {
    const subj = (kind === "ok" ? "✅ " : "⚠️ ") + "Wabi Navi foto: " + (msg.getSubject() || "");
    GmailApp.sendEmail(to, subj, body);
  } catch (e) { Logger.log("replyPhoto: " + e); }
}

// ── スレッドを処理済みにする（ラベル付与＋ゴミ箱へ） ─────────────
function markDone_(thread, label) {
  try {
    thread.addLabel(label);
    thread.moveToTrash();   // 受信箱が溜まらないようゴミ箱へ
  } catch (e) { Logger.log("markDone: " + e); }
}

// ── メールFromから素のアドレスを抜き出す ─────────────────────────
function extractEmail_(from) {
  const m = (from || "").match(/<([^>]+)>/);
  return (m ? m[1] : from).trim();
}


// ════════════════════════════════════════════════════════════════════
//  📣 SNS投稿パッケージ生成 / Pakiet do social media（v7.3）
//
//  プロモ写真（件名 promo_○○）を受け取ったとき、写真をGitHubに保存する
//  のに加えて、Facebook / Instagram / TikTok / WhatsApp 向けの投稿文を
//  自動生成し、送信者へメールで届ける。
//
//  方針：
//   ・ポーランド語がメイン、英語は補助（短く下に添える）
//   ・完全自動投稿ではなく「コピペで貼るだけ」の半自動（壊れない・無料・審査不要）
//   ・メール本文に書いた宣伝文があれば、それを各SNS向けに整形して使う
//     （空なら定型のテンプレートを使う）
//
//  将来 Facebook/Instagram の自動投稿（C案）へ移行する際も、
//  ここで作る文面（buildSocialText_）はそのまま再利用できる。
// ════════════════════════════════════════════════════════════════════

// 店の固定情報（必要なら書き換え）
const SNS_INFO = {
  name:  "Wabi Navi",
  addr:  "ul. Małe Garbary 5A, Toruń",
  tags:  "#WabiNavi #ramen #Toruń #kuchniajapońska #ramentoruń #japanesefood",
  igtags:"#WabiNavi #ramen #Toruń #kuchniajapońska #japanesefood #foodie #torunfood #ramenlover #japanindorf #instafood"
};

// ── メール本文から宣伝文（キャプション）を取り出す ───────────────
//   引用（>）や署名らしき行を除き、最初のまとまった文章を使う。
function extractCaption_(body) {
  if (!body) return "";
  const lines = body.split(/\r?\n/)
    .map(s => s.trim())
    .filter(s => s && s.indexOf(">") !== 0 &&
                 !/^(sent from|wysłane z|--|__)/i.test(s));
  // 最初の空行までをキャプションとみなす（最大3行）
  const out = [];
  for (const l of lines) {
    if (out.length >= 3) break;
    out.push(l);
  }
  return out.join(" ").slice(0, 300);
}

// ── 各SNS向けの本文を組み立て（PL主体・EN補助） ─────────────────
function buildSocialText_(caption) {
  const plCore = caption ||
    "Nowość w Wabi Navi! Zapraszamy na pyszne, japońskie smaki. 🍜";
  const enCore = caption
    ? "" // 独自キャプションがある時は英語補助を短く付ける
    : "New at Wabi Navi! Come and taste authentic Japanese flavours. 🍜";

  const en = enCore ? ("\n\n— EN —\n" + enCore) : "";

  // Facebook：長め本文OK・住所や案内を入れる
  const facebook =
    plCore + "\n\n📍 " + SNS_INFO.name + ", " + SNS_INFO.addr +
    "\n🍽 Rezerwacje / zamówienia: menu na stronie" +
    (enCore ? "\n\n— EN —\n" + enCore + "\n📍 " + SNS_INFO.addr : "") +
    "\n\n" + SNS_INFO.tags;

  // Instagram：ハッシュタグ多め・キャプション＋タグ
  const instagram =
    plCore +
    (enCore ? "\n\n" + enCore : "") +
    "\n\n📍 " + SNS_INFO.addr +
    "\n.\n.\n" + SNS_INFO.igtags;

  // TikTok：短くフック重視・タグ少なめ
  const tiktok =
    (caption ? caption : "Japoński ramen w Toruniu! 🍜🔥") +
    "\n" + "#WabiNavi #ramen #Toruń #fyp #foodtiktok";

  // WhatsApp：会員・常連への一斉配信向け（絵文字＋短め＋来店誘導）
  const whatsapp =
    "🍜 *" + SNS_INFO.name + "*\n" +
    plCore +
    (enCore ? "\n\n" + enCore : "") +
    "\n\n📍 " + SNS_INFO.addr +
    "\nZapraszamy! / Welcome!";

  return { facebook, instagram, tiktok, whatsapp, plCore, enCore };
}

// ── 投稿パッケージをメールで送信 ─────────────────────────────────
function sendSocialPackage_(c, to, path, caption, attachment) {
  const t = buildSocialText_(caption);
  const imgUrl = "https://" + c.owner.toLowerCase() + ".github.io/" +
                 c.repo + "/" + path;   // 例：…github.io/wabinavi-order/images/promo/lato.jpg

  const sep = "\n────────────────────\n";
  const body =
    "📣 PAKIET DO SOCIAL MEDIA / SNS投稿パッケージ\n" +
    "Skopiuj i wklej do wybranej aplikacji. / コピーして各SNSに貼り付けてください。\n" +
    sep +
    "■ FACEBOOK\n" + t.facebook + "\n" +
    sep +
    "■ INSTAGRAM\n" + t.instagram + "\n" +
    sep +
    "■ TIKTOK\n" + t.tiktok + "\n" +
    sep +
    "■ WHATSAPP\n" + t.whatsapp + "\n" +
    sep +
    "🖼 Zdjęcie / 写真：\n" +
    "・W załączniku (do FB/IG/TikTok/WhatsApp) / 添付ファイルを使用\n" +
    "・Online: " + imgUrl + "\n" +
    sep +
    "ℹ️ To NIE jest automatyczny post — wklej ręcznie. / これは自動投稿ではありません。手で貼り付けてください。\n" +
    "Zdjęcie zostało już zapisane na stronie i w aplikacji. / 写真はHP・アプリには保存済みです。";

  const opts = { name: "Wabi Navi Foto" };
  try {
    if (attachment) opts.attachments = [attachment.copyBlob()];
  } catch (e) {}

  GmailApp.sendEmail(to,
    "📣 SNS投稿パッケージ / Pakiet social media — " + path.split("/").pop(),
    body, opts);
}


// ════════════════════════════════════════════════════════════════════
//  🍳 キッチン表示システム / Kitchen Display System (KDS) — v8.0
//
//  タブレット用の kitchen.html と組み合わせて使う。
//   ・卓ごとのカードで注文を表示（先頭が大きく、次の卓以降も同一画面に）
//   ・だし別のラーメン杯数・替え玉の玉数・餃子/揚げ物の皿数をまとめて集計
//   ・10分で黄色点滅、20分で赤色点滅
//   ・調理済み（GOTOWE）にすると自動で非表示、履歴タブから差し戻し可能
//
//  状態は注文シートR列（STATUS）で管理する。既存のプルダウンをそのまま使用：
//     NOWE / NEW              … 未処理（キッチンに表示）
//     W TRAKCIE / IN PROGRESS … 調理中（キッチンに表示）
//     GOTOWE / READY          … 調理済み・配膳OK（キッチンから消える）
//     WYDANE / SERVED         … 提供済み（キッチンから消える）
// ════════════════════════════════════════════════════════════════════

const KDS_MAX_ROWS = 1500;   // 直近何行まで読むか（速度対策）


// ── 注文行を「新しい順」で記録する ────────────────────────────────
//    appendRow（末尾に追加）をやめ、2行目に差し込む。
//    これで一番新しい注文が常にヘッダーのすぐ下に来る。
function prependOrderRows_(sh, rows) {
  if (!rows || !rows.length) return 2;
  const width = rows[0].length;
  sh.insertRowsBefore(2, rows.length);
  const rng = sh.getRange(2, 1, rows.length, width);
  rng.clearFormat();          // 挿入行はヘッダーの濃色書式を引き継ぐため必ずリセット
  rng.setValues(rows);
  return 2;
}


// ── キッチン用の列を用意（初回1回だけ実行）──────────────────────
//    T列＝商品ID。注文アプリ v42 以降が送信する。
//    これでキッチン画面がラーメンのだしベースや商品分類を判別できる。
function setupKitchenColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) { ui.alert("注文シートがありません / Brak arkusza zamówień"); return; }
  sh.getRange(1, 20).setValue("商品ID / Item ID")
    .setFontWeight("bold").setBackground("#2c1810").setFontColor("#c8a96e")
    .setHorizontalAlignment("center").setWrap(true).setFontSize(9);
  sh.setColumnWidth(20, 110);
  ui.alert("✅ キッチン用の列（T列＝商品ID）を追加しました",
    "これ以降の注文には商品IDが記録され、キッチン画面で\n" +
    "「だし別のラーメン杯数」「餃子の皿数」などの集計ができます。\n\n" +
    "※過去の注文には商品IDが入りませんが、商品名で照合するので\n" +
    "　キッチン画面には問題なく表示されます。",
    ui.ButtonSet.OK);
}


// ── キッチン画面へ渡すデータ ──────────────────────────────────────
//    GET ?action=getKitchen&password=…&mode=active|history&date=yyyy-mm-dd
function getKitchenData(password, mode, dateStr) {
  const auth = verifyStaffPassword(password);
  if (auth !== "ok") return { status: "unauthorized", message: STAFF_AUTH_MSG[auth] };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ORDERS);
  const now = new Date();
  const today = dateStr || Utilities.formatDate(now, "Europe/Warsaw", "yyyy-MM-dd");
  const empty = { status: "ok", mode: mode || "active", today: today,
                  serverTime: now.getTime(), orders: [] };
  if (!sh) return empty;

  const last = sh.getLastRow();
  if (last < 2) return empty;
  const width = Math.min(Math.max(sh.getLastColumn(), 20), sh.getMaxColumns());
  const n = Math.min(last - 1, KDS_MAX_ROWS);
  const data = sh.getRange(2, 1, n, width).getValues();

  const wantHistory = (mode === "history");
  const meta = {};      // orderId → 注文の共通情報
  const seq  = [];      // 出現順

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const oid = (r[0] || "").toString().replace(/^⚠️/, "").trim();
    if (!oid) continue;
    if (oid.indexOf("💳") === 0) continue;        // 会計リクエスト行は調理対象外
    const t = r[1];
    if (!(t instanceof Date)) continue;
    if (Utilities.formatDate(t, "Europe/Warsaw", "yyyy-MM-dd") !== today) continue;

    if (!meta[oid]) {
      const tableTxt = (r[2] || "").toString();
      const tm = tableTxt.match(/(\d+)/);
      meta[oid] = {
        orderId: oid,
        ts: t.getTime(),
        hhmm: Utilities.formatDate(t, "Europe/Warsaw", "HH:mm"),
        table: tm ? Number(tm[1]) : 0,
        tableTxt: tableTxt,
        guests: (r[3] || "").toString(),
        uwagi: "", member: "", memberName: "",
        items: [], doneCount: 0, totalCount: 0
      };
      seq.push(oid);
    }
    const o = meta[oid];
    // 備考・会員名は1品目の行にしか入らないため、空でない値を拾って保持する
    if (r[11]) o.uwagi = (r[11] || "").toString();
    if (r[12]) o.member = (r[12] || "").toString();
    if (r[13]) o.memberName = (r[13] || "").toString();

    const st   = (r[17] || "").toString();
    const done = /GOTOWE|WYDANE|READY|SERVED/i.test(st);
    o.totalCount++;
    if (done) o.doneCount++;
    if (done !== wantHistory) continue;           // 表示モードに合う品だけ入れる

    o.items.push({
      no:      Number(r[4]) || 0,
      itemId:  (r[19] || "").toString().trim(),
      namePL:  (r[5]  || "").toString(),
      nameEN:  (r[6]  || "").toString(),
      qty:     Number(r[7]) || 1,
      memo:    (r[10] || "").toString(),
      status:  st
    });
  }

  const list = seq.map(id => meta[id]).filter(o => o.items.length > 0);
  // 調理中は古い順（先に入った注文を先に作る）。履歴は新しい順。
  list.sort((a, b) => wantHistory ? (b.ts - a.ts) : (a.ts - b.ts));

  return { status: "ok", mode: mode || "active", today: today,
           serverTime: now.getTime(), orders: list };
}


// ── キッチンからの状態更新 ────────────────────────────────────────
//    POST { action:"kitchenUpdate", password, orderId,
//           itemNo（省略＝その注文の全品）, newStatus }
//    行番号ではなく「注文ID＋品番」で特定する。新着が上に差し込まれて
//    行番号がずれても、確実に正しい品を更新できる。
function handleKitchenUpdate(data) {
  const auth = verifyStaffPassword(data.password);
  if (auth !== "ok") return jsonResponse({ status: "unauthorized", message: STAFF_AUTH_MSG[auth] });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) return jsonResponse({ status: "error", message: "注文シートがありません" });

  const oid = (data.orderId || "").toString().trim();
  if (!oid) return jsonResponse({ status: "error", message: "orderId required" });
  const hasItemNo = (data.itemNo !== undefined && data.itemNo !== null && data.itemNo !== "");
  const itemNo = hasItemNo ? Number(data.itemNo) : null;

  const ALLOWED = ["NOWE / NEW", "W TRAKCIE / IN PROGRESS",
                   "GOTOWE / READY", "WYDANE / SERVED"];
  const st = (data.newStatus || "GOTOWE / READY").toString();
  if (ALLOWED.indexOf(st) === -1) {
    return jsonResponse({ status: "error", message: "Nieznany status / 不正なステータス" });
  }

  const last = sh.getLastRow();
  if (last < 2) return jsonResponse({ status: "notFound" });
  const rows = sh.getRange(2, 1, last - 1, 5).getValues();   // A〜E列

  let updated = 0;
  for (let i = 0; i < rows.length; i++) {
    const v = (rows[i][0] || "").toString().replace(/^⚠️/, "").trim();
    if (v !== oid) continue;
    if (itemNo !== null && Number(rows[i][4]) !== itemNo) continue;
    sh.getRange(i + 2, 18).setValue(st);
    updated++;
  }
  if (!updated) return jsonResponse({ status: "notFound", message: "該当の注文が見つかりません" });
  return jsonResponse({ status: "ok", updated: updated, newStatus: st });
}


// ── 提供済みの行を薄く表示（任意・メニューから手動実行）──────────
//    行は削除しないので、あとから上にスクロールして確認できる。
function dimServedOrderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) { ui.alert("注文シートがありません"); return; }
  const last = sh.getLastRow();
  if (last < 2) { ui.alert("注文がありません"); return; }
  const n = Math.min(last - 1, KDS_MAX_ROWS);
  const st = sh.getRange(2, 18, n, 1).getValues();
  const w  = Math.min(20, sh.getMaxColumns());
  let dimmed = 0;
  for (let i = 0; i < n; i++) {
    if (/GOTOWE|WYDANE|READY|SERVED/i.test((st[i][0] || "").toString())) {
      sh.getRange(i + 2, 1, 1, w).setFontColor("#b0aaa4");
      dimmed++;
    }
  }
  ui.alert("✅ 完了済みの表示を薄くしました",
    dimmed + " 行をグレーにしました。\n" +
    "行は削除していないので、いつでも確認できます。",
    ui.ButtonSet.OK);
}
