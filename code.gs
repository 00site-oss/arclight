/**
 * Arclight — Google Apps Script Backend
 * Deploy as Web App: Execute as Me, Access: Anyone
 *
 * After deploy, run setupAdmin_() once from the editor.
 */

const SHEETS = {
  USERS: 'Users',
  TRANSACTIONS: 'Transactions',
  OFFERS: 'Offers',
  QUIZ_QUESTIONS: 'QuizQuestions',
  QUIZ_ATTEMPTS: 'QuizAttempts',
  COMPLETIONS: 'Completions',
  CASHOUTS: 'Cashouts',
  POLLS: 'Polls',
  VOTES: 'Votes',
  CHAT: 'Chat',
  NEWS: 'News',
  CODES: 'Codes',
  SETTINGS: 'Settings',
  RANKS: 'Ranks',
  REFERRALS: 'Referrals',
  CASH_METHODS: 'CashMethods'
};

const USERNAME_CHANGE_COST = 250;
const ADMIN_TOKEN_SALT = 'arclight_admin_v1';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    ensureSheets_();
    const action = (e.parameter && e.parameter.action) ||
      (e.postData && JSON.parse(e.postData.contents).action) || '';
    const body = e.postData ? JSON.parse(e.postData.contents) : (e.parameter || {});
    const data = Object.assign({}, e.parameter || {}, body);

    let result;
    switch (action) {
      case 'init': result = initSystem_(); break;
      case 'login': result = login_(data); break;
      case 'register': result = register_(data); break;
      case 'getUser': result = getUser_(data); break;
      case 'updateProfile': result = updateProfile_(data); break;
      case 'changeUsername': result = changeUsername_(data); break;
      case 'updateAccount': result = updateAccount_(data); break;
      case 'getPartners': result = getPartners_(data); break;
      case 'getOffer': result = getOffer_(data); break;
      case 'getQuiz': result = getQuiz_(data); break;
      case 'submitQuiz': result = submitQuiz_(data); break;
      case 'getBalance': result = getBalance_(data); break;
      case 'getCredits': result = getRecentCredits_(); break;
      case 'getNews': result = getNews_(); break;
      case 'getPoll': result = getDailyPoll_(); break;
      case 'votePoll': result = votePoll_(data); break;
      case 'getGoal': result = getCurrentGoal_(data); break;
      case 'getCashout6': result = getCashout6_(data); break;
      case 'claimCashout6': result = claimCashout6_(data); break;
      case 'redeemCode': result = redeemCode_(data); break;
      case 'requestCashout': result = requestCashout_(data); break;
      case 'getHistory': result = getHistory_(data); break;
      case 'getChat': result = getChat_(data); break;
      case 'sendChat': result = sendChat_(data); break;
      case 'getRanks': result = getRanks_(); break;
      case 'getHome': result = getHome_(data); break;
      case 'creditOffer': result = creditOffer_(data); break;
      case 'getCashMethods': result = getCashMethods_(); break;
      case 'adminLogin': result = adminLogin_(data); break;
      case 'adminStats': result = adminStats_(data); break;
      case 'adminListUsers': result = adminListUsers_(data); break;
      case 'adminUpdateUser': result = adminUpdateUser_(data); break;
      case 'adminListOffers': result = adminListOffers_(data); break;
      case 'adminSaveOffer': result = adminSaveOffer_(data); break;
      case 'adminDeleteOffer': result = adminDeleteOffer_(data); break;
      case 'adminListQuestions': result = adminListQuestions_(data); break;
      case 'adminSaveQuestion': result = adminSaveQuestion_(data); break;
      case 'adminDeleteQuestion': result = adminDeleteQuestion_(data); break;
      case 'adminListCashouts': result = adminListCashouts_(data); break;
      case 'adminProcessCashout': result = adminProcessCashout_(data); break;
      case 'adminListPolls': result = adminListPolls_(data); break;
      case 'adminSavePoll': result = adminSavePoll_(data); break;
      case 'adminListNews': result = adminListNews_(data); break;
      case 'adminSaveNews': result = adminSaveNews_(data); break;
      case 'adminDeleteNews': result = adminDeleteNews_(data); break;
      case 'adminListCodes': result = adminListCodes_(data); break;
      case 'adminSaveCode': result = adminSaveCode_(data); break;
      case 'adminListRanks': result = adminListRanks_(data); break;
      case 'adminSaveRank': result = adminSaveRank_(data); break;
      case 'adminGetSettings': result = adminGetSettings_(data); break;
      case 'adminSetSetting': result = adminSetSetting_(data); break;
      case 'adminListChat': result = adminListChat_(data); break;
      case 'adminDeleteChat': result = adminDeleteChat_(data); break;
      case 'adminListCashMethods': result = adminListCashMethods_(data); break;
      case 'adminSaveCashMethod': result = adminSaveCashMethod_(data); break;
      case 'adminDeleteCashMethod': result = adminDeleteCashMethod_(data); break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.SETTINGS)) initSystem_();
}

function initSystem_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheet_(ss, SHEETS.USERS, [
    'id', 'username', 'email', 'password_hash', 'balance', 'on_hold',
    'total_earned', 'total_redeemed', 'rank', 'country', 'ref_source',
    'referrer_id', 'avatar', 'status', 'created_at', 'updated_at'
  ]);
  createSheet_(ss, SHEETS.TRANSACTIONS, [
    'id', 'user_id', 'type', 'amount', 'balance_after', 'description', 'ref_id', 'created_at'
  ]);
  createSheet_(ss, SHEETS.OFFERS, [
    'id', 'name', 'partner', 'description', 'points', 'daily_limit',
    'pass_percent', 'max_tries', 'active', 'url', 'created_at'
  ]);
  createSheet_(ss, SHEETS.QUIZ_QUESTIONS, [
    'id', 'offer_id', 'question', 'options', 'correct_index', 'order'
  ]);
  createSheet_(ss, SHEETS.QUIZ_ATTEMPTS, [
    'id', 'user_id', 'offer_id', 'score_percent', 'passed', 'points_awarded',
    'answers', 'try_number', 'created_at'
  ]);
  createSheet_(ss, SHEETS.COMPLETIONS, [
    'id', 'user_id', 'offer_id', 'partner', 'points', 'status', 'created_at', 'approved_at'
  ]);
  createSheet_(ss, SHEETS.CASHOUTS, [
    'id', 'user_id', 'method', 'amount', 'points', 'status', 'details', 'created_at', 'processed_at'
  ]);
  createSheet_(ss, SHEETS.POLLS, [
    'id', 'question', 'options', 'reward', 'active', 'created_at', 'ends_at'
  ]);
  createSheet_(ss, SHEETS.VOTES, [
    'id', 'poll_id', 'user_id', 'option_index', 'created_at'
  ]);
  createSheet_(ss, SHEETS.CHAT, [
    'id', 'user_id', 'username', 'badge', 'message', 'deleted', 'deleted_by', 'created_at'
  ]);
  createSheet_(ss, SHEETS.NEWS, [
    'id', 'title', 'body', 'created_at'
  ]);
  createSheet_(ss, SHEETS.CODES, [
    'id', 'code', 'points', 'max_uses', 'uses', 'active', 'expires_at', 'created_at'
  ]);
  createSheet_(ss, SHEETS.SETTINGS, ['key', 'value']);
  createSheet_(ss, SHEETS.RANKS, ['rank', 'earnings', 'referrals', 'rewards', 'perks']);
  createSheet_(ss, SHEETS.REFERRALS, ['id', 'referrer_id', 'referred_id', 'active', 'created_at']);
  createSheet_(ss, SHEETS.CASH_METHODS, ['id', 'name', 'min_points', 'icon', 'active', 'sort_order']);
  seedData_();
  setupAdmin_();
  return { ok: true, message: 'All sheets created and seeded.' };
}

function createSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function seedData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (settings.getLastRow() <= 1) {
    [
      ['username_change_cost', String(USERNAME_CHANGE_COST)],
      ['min_cashout', '300'],
      ['site_name', 'Arclight'],
      ['cashout6_slots', '6']
    ].forEach(r => settings.appendRow(r));
  }

  const ranks = ss.getSheetByName(SHEETS.RANKS);
  if (ranks.getLastRow() <= 1) {
    [
      ['ROOKIE', 0, 0, 0, ''],
      ['JUNIOR', 500, 10, 1, 'Receive Points'],
      ['SENIOR', 3000, 50, 2, 'Instant Rewards;Lower Cashout Minimum'],
      ['VETERAN', 10000, 100, 3, 'Chat Helper Jr.'],
      ['ELITE', 35000, 250, 4, 'Chat Helper Sr.'],
      ['LEGEND', 110000, 500, 5, 'Earnings Privacy'],
      ['PRIME', 260000, 1000, 6, '']
    ].forEach(r => ranks.appendRow(r));
  }

  const offers = ss.getSheetByName(SHEETS.OFFERS);
  if (offers.getLastRow() <= 1) {
    const partners = [
      ['Revenue Universe', 'Revenue Universe', 'Complete surveys and offers. Limited to 3 completes a day.', 50, 3, 70, 3],
      ['WallAds', 'WallAds', 'Offers and surveys. Limited to 5 completes a day.', 40, 5, 60, 3],
      ['Torox', 'Torox', 'Games and multiple earn methods.', 30, 0, 50, 5],
      ['lootably', 'lootably', 'Offers, surveys and watch content. Limited to 4 completes a day.', 45, 4, 70, 3],
      ['ayeT-Studios', 'ayeT-Studios', 'Offers and surveys from ayeT-Studios.', 35, 0, 60, 3],
      ['Monlix', 'Monlix', 'Surveys and PTC. Limited to 5 completes a day.', 25, 5, 50, 4],
      ['Wannads', 'Wannads', 'Surveys. Limited to 3 completes a day.', 50, 3, 70, 3],
      ['Adwall', 'Adwall', 'Offers and surveys.', 20, 0, 50, 5]
    ];
    partners.forEach((p, i) => {
      offers.appendRow(['off_' + (i + 1), p[0], p[1], p[2], p[3], p[4], p[5], p[6], true, '', new Date().toISOString()]);
    });
  }

  const quizQ = ss.getSheetByName(SHEETS.QUIZ_QUESTIONS);
  if (quizQ.getLastRow() <= 1) {
    const questions = [
      ['off_1', 'What is a survey offerwall primarily used for?', JSON.stringify(['Playing games only', 'Completing surveys and offers for rewards', 'Streaming music', 'Sending emails']), 1, 1],
      ['off_1', 'How many completes per day does Revenue Universe allow?', JSON.stringify(['1', '2', '3', 'Unlimited']), 2, 2],
      ['off_1', 'Points are typically awarded when?', JSON.stringify(['Before starting', 'After the offer is completed and verified', 'Never', 'Only on weekends']), 1, 3],
      ['off_1', 'Which action can disqualify an offer?', JSON.stringify(['Finishing all steps', 'Using VPN or fake information', 'Reading instructions', 'Waiting for credit']), 1, 4],
      ['off_2', 'WallAds offers come from which network?', JSON.stringify(['Google', 'PointClickTrack', 'Facebook', 'Amazon']), 1, 1],
      ['off_2', 'What is the daily complete limit for WallAds?', JSON.stringify(['3', '5', '10', 'Unlimited']), 1, 2],
      ['off_2', 'Failing a quiz below the pass rate results in?', JSON.stringify(['Full points', 'No change', 'Half points as a penalty', 'Double points']), 2, 3],
      ['off_3', 'Torox is best known for?', JSON.stringify(['Only surveys', 'Games and multiple earn methods', 'Only PTC', 'Crypto mining']), 1, 1],
      ['off_3', 'What percent do you need to pass a quiz?', JSON.stringify(['The offer pass_percent setting', 'Always 100%', 'Always 0%', 'Random']), 0, 2]
    ];
    questions.forEach((q, i) => quizQ.appendRow(['q_' + (i + 1), q[0], q[1], q[2], q[3], q[4]]));
  }

  const polls = ss.getSheetByName(SHEETS.POLLS);
  if (polls.getLastRow() <= 1) {
    polls.appendRow([
      'poll_1',
      'How did you hear about Arclight?',
      JSON.stringify(['Friend', 'Search', 'Social media', 'Other']),
      1, true, new Date().toISOString(), ''
    ]);
  }

  const news = ss.getSheetByName(SHEETS.NEWS);
  if (news.getLastRow() <= 1) {
    [
      ['Welcome to Arclight', 'Your independent earn station is online.'],
      ['Cashouts available', 'PayPal and crypto payouts are supported when you hit the minimum.'],
      ['Quiz tips', 'Read each question carefully — pass rate varies by offer.']
    ].forEach((n, i) => news.appendRow(['news_' + (i + 1), n[0], n[1], new Date().toISOString()]));
  }

  const codes = ss.getSheetByName(SHEETS.CODES);
  if (codes.getLastRow() <= 1) {
    codes.appendRow(['code_1', 'WELCOME50', 50, 100, 0, true, '', new Date().toISOString()]);
    codes.appendRow(['code_2', 'BONUS100', 100, 50, 0, true, '', new Date().toISOString()]);
  }

  const users = ss.getSheetByName(SHEETS.USERS);
  if (users.getLastRow() <= 1) {
    users.appendRow([
      'usr_iml', 'IML', 'iml@example.com', hash_('password123'),
      1, 200, 10117.99, 8200, 'VETERAN', 'United States', 'none', '',
      '', 'offline', '2022-01-06T00:00:00.000Z', new Date().toISOString()
    ]);
  }

  const cm = ss.getSheetByName(SHEETS.CASH_METHODS);
  if (cm && cm.getLastRow() <= 1) {
    [
      ['PayPal', 'PayPal', 300, 'PP', true, 1],
      ['Bitcoin', 'Bitcoin', 500, '₿', true, 2],
      ['Litecoin', 'Litecoin', 300, 'Ł', true, 3],
      ['Gift Card', 'Gift Card', 500, 'GC', true, 4]
    ].forEach(r => cm.appendRow(r));
  }
}

function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function rows_(name) {
  const sh = sheet_(name);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findUser_(idOrName) {
  const users = rows_(SHEETS.USERS);
  return users.find(u =>
    String(u.id) === String(idOrName) ||
    String(u.username).toLowerCase() === String(idOrName).toLowerCase()
  ) || null;
}

function userRowIndex_(userId) {
  const sh = sheet_(SHEETS.USERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) return i + 1;
  }
  return -1;
}

function hash_(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

function uid_(prefix) {
  return prefix + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function addTransaction_(userId, type, amount, description, refId) {
  const user = findUser_(userId);
  if (!user) throw new Error('User not found');
  const newBal = Number(user.balance) + Number(amount);
  sheet_(SHEETS.TRANSACTIONS).appendRow([
    uid_('txn'), userId, type, amount, newBal, description || '', refId || '', new Date().toISOString()
  ]);
  const row = userRowIndex_(userId);
  if (row > 0) {
    sheet_(SHEETS.USERS).getRange(row, 5).setValue(newBal);
    if (amount > 0) {
      sheet_(SHEETS.USERS).getRange(row, 7).setValue(Number(user.total_earned) + Number(amount));
    }
    sheet_(SHEETS.USERS).getRange(row, 16).setValue(new Date().toISOString());
  }
  return newBal;
}

function calcRank_(totalEarned, referrals) {
  const ranks = rows_(SHEETS.RANKS).slice().reverse();
  for (const r of ranks) {
    if (Number(totalEarned) >= Number(r.earnings) || Number(referrals) >= Number(r.referrals)) return r.rank;
  }
  return 'ROOKIE';
}

function register_(data) {
  const username = String(data.username || '').trim();
  const email = String(data.email || '').trim().toLowerCase();
  const password = String(data.password || '');
  if (!username || !email || !password) return { ok: false, error: 'Username, email, and password required.' };
  if (findUser_(username)) return { ok: false, error: 'Username already taken.' };
  if (rows_(SHEETS.USERS).find(u => String(u.email).toLowerCase() === email)) {
    return { ok: false, error: 'Email already registered.' };
  }
  const id = uid_('usr');
  sheet_(SHEETS.USERS).appendRow([
    id, username, email, hash_(password), 0, 0, 0, 0, 'ROOKIE',
    data.country || '', data.ref_source || 'none', data.referrer_id || '',
    '', 'online', new Date().toISOString(), new Date().toISOString()
  ]);
  if (data.referrer_id) {
    sheet_(SHEETS.REFERRALS).appendRow([uid_('ref'), data.referrer_id, id, false, new Date().toISOString()]);
  }
  return { ok: true, user: getUser_({ user_id: id }).user };
}

function login_(data) {
  const user = findUser_(data.username || data.email);
  if (!user) return { ok: false, error: 'User not found.' };
  if (hash_(String(data.password || '')) !== String(user.password_hash)) {
    return { ok: false, error: 'Invalid password.' };
  }
  const row = userRowIndex_(user.id);
  if (row > 0) sheet_(SHEETS.USERS).getRange(row, 14).setValue('online');
  return { ok: true, user: sanitizeUser_(user) };
}

function sanitizeUser_(u) {
  return {
    id: u.id, username: u.username, email: u.email,
    balance: Number(u.balance) || 0, on_hold: Number(u.on_hold) || 0,
    total_earned: Number(u.total_earned) || 0, total_redeemed: Number(u.total_redeemed) || 0,
    rank: u.rank || 'ROOKIE', country: u.country || '', ref_source: u.ref_source || 'none',
    status: u.status || 'offline', created_at: u.created_at,
    contest_winnings: 0, lotteries: 0, referrals: countReferrals_(u.id),
    forum_posts: 0, trades_up: 0, trades_down: 0
  };
}

function countReferrals_(userId) {
  return rows_(SHEETS.REFERRALS).filter(r => String(r.referrer_id) === String(userId)).length;
}

function getUser_(data) {
  const user = findUser_(data.user_id || data.username || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  return { ok: true, user: sanitizeUser_(user) };
}

function getBalance_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  return { ok: true, balance: Number(user.balance) || 0, on_hold: Number(user.on_hold) || 0 };
}

function changeUsername_(data) {
  const user = findUser_(data.user_id || data.username || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  const newName = String(data.new_username || '').trim();
  if (!newName) return { ok: false, error: 'New username required.' };
  if (newName.length < 3 || newName.length > 20) return { ok: false, error: 'Username must be 3–20 characters.' };
  if (!/^[a-zA-Z0-9_]+$/.test(newName)) return { ok: false, error: 'Username can only contain letters, numbers, and underscores.' };
  if (String(user.username).toLowerCase() === newName.toLowerCase()) return { ok: false, error: 'That is already your username.' };
  if (findUser_(newName)) return { ok: false, error: 'Username already taken.' };
  const cost = USERNAME_CHANGE_COST;
  const balance = Number(user.balance) || 0;
  if (balance < cost) {
    return { ok: false, error: 'Not enough points. Username change costs ' + cost + ' PTS. You have ' + balance + ' PTS.', cost: cost, balance: balance };
  }
  if (data.confirm !== true && data.confirm !== 'true') {
    return { ok: true, requires_confirm: true, cost: cost, balance: balance, message: 'Changing your username costs ' + cost + ' PTS. Proceed?' };
  }
  const newBal = addTransaction_(user.id, 'username_change', -cost, 'Username change: ' + user.username + ' → ' + newName, '');
  const row = userRowIndex_(user.id);
  sheet_(SHEETS.USERS).getRange(row, 2).setValue(newName);
  sheet_(SHEETS.USERS).getRange(row, 16).setValue(new Date().toISOString());
  return { ok: true, username: newName, balance: newBal, cost: cost, message: 'Username changed to ' + newName + '. ' + cost + ' PTS deducted.' };
}

function updateAccount_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  const row = userRowIndex_(user.id);
  const sh = sheet_(SHEETS.USERS);
  if (data.email) {
    const email = String(data.email).trim().toLowerCase();
    if (!email.includes('@')) return { ok: false, error: 'Invalid email.' };
    sh.getRange(row, 3).setValue(email);
  }
  if (data.password) {
    if (String(data.password).length < 6) return { ok: false, error: 'Password must be at least 6 characters.' };
    sh.getRange(row, 4).setValue(hash_(String(data.password)));
  }
  sh.getRange(row, 16).setValue(new Date().toISOString());
  return { ok: true, message: 'Account updated.', user: getUser_({ user_id: user.id }).user };
}

function updateProfile_(data) { return updateAccount_(data); }

function getPartners_(data) {
  const userId = data.user_id || '';
  const offers = rows_(SHEETS.OFFERS).filter(o => o.active === true || o.active === 'TRUE' || o.active === '');
  let attemptsByOffer = {};
  if (userId) {
    rows_(SHEETS.QUIZ_ATTEMPTS).filter(a => String(a.user_id) === String(userId)).forEach(a => {
      const k = String(a.offer_id);
      attemptsByOffer[k] = (attemptsByOffer[k] || 0) + 1;
    });
  }
  return {
    ok: true,
    partners: offers.map(o => {
      const tries = attemptsByOffer[String(o.id)] || 0;
      const maxTries = Number(o.max_tries) || 3;
      return {
        id: o.id, name: o.name, partner: o.partner, description: o.description,
        points: Number(o.points) || 0, daily_limit: Number(o.daily_limit) || 0,
        pass_percent: Number(o.pass_percent) || 70, max_tries: maxTries,
        tries_used: tries, tries_left: Math.max(0, maxTries - tries), has_quiz: true
      };
    })
  };
}

function getOffer_(data) {
  const offer = rows_(SHEETS.OFFERS).find(o => String(o.id) === String(data.offer_id));
  if (!offer) return { ok: false, error: 'Offer not found.' };
  return {
    ok: true,
    offer: {
      id: offer.id, name: offer.name, partner: offer.partner, description: offer.description,
      points: Number(offer.points) || 0, pass_percent: Number(offer.pass_percent) || 70,
      max_tries: Number(offer.max_tries) || 3
    }
  };
}

function getQuiz_(data) {
  const offerId = data.offer_id;
  if (!offerId) return { ok: false, error: 'offer_id required.' };
  const offer = rows_(SHEETS.OFFERS).find(o => String(o.id) === String(offerId));
  if (!offer) return { ok: false, error: 'Offer not found.' };
  const user = findUser_(data.user_id || data.username || 'IML');
  const triesUsed = user
    ? rows_(SHEETS.QUIZ_ATTEMPTS).filter(a => String(a.user_id) === String(user.id) && String(a.offer_id) === String(offerId)).length
    : 0;
  const maxTries = Number(offer.max_tries) || 3;
  if (triesUsed >= maxTries) {
    return { ok: false, error: 'No tries left for this offer.', tries_used: triesUsed, max_tries: maxTries };
  }
  const questions = rows_(SHEETS.QUIZ_QUESTIONS)
    .filter(q => String(q.offer_id) === String(offerId))
    .sort((a, b) => Number(a.order) - Number(b.order))
    .map(q => {
      let options = q.options;
      try { options = JSON.parse(q.options); } catch (e) {}
      return { id: q.id, question: q.question, options: options };
    });
  if (!questions.length) return { ok: false, error: 'No quiz questions for this offer yet.' };
  return {
    ok: true,
    offer: {
      id: offer.id, name: offer.name, points: Number(offer.points) || 0,
      pass_percent: Number(offer.pass_percent) || 70, max_tries: maxTries,
      tries_used: triesUsed, tries_left: maxTries - triesUsed
    },
    questions: questions
  };
}

function submitQuiz_(data) {
  const offerId = data.offer_id;
  const user = findUser_(data.user_id || data.username || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  if (!offerId) return { ok: false, error: 'offer_id required.' };
  const offer = rows_(SHEETS.OFFERS).find(o => String(o.id) === String(offerId));
  if (!offer) return { ok: false, error: 'Offer not found.' };
  const maxTries = Number(offer.max_tries) || 3;
  const prior = rows_(SHEETS.QUIZ_ATTEMPTS).filter(a => String(a.user_id) === String(user.id) && String(a.offer_id) === String(offerId));
  if (prior.length >= maxTries) {
    return { ok: false, error: 'No tries left for this offer.', tries_used: prior.length, max_tries: maxTries };
  }
  const questions = rows_(SHEETS.QUIZ_QUESTIONS).filter(q => String(q.offer_id) === String(offerId));
  if (!questions.length) return { ok: false, error: 'No questions found.' };
  let answers = data.answers;
  if (typeof answers === 'string') { try { answers = JSON.parse(answers); } catch (e) { answers = {}; } }
  answers = answers || {};
  let correct = 0;
  const total = questions.length;
  const details = [];
  questions.forEach(q => {
    const selected = answers[q.id];
    const isCorrect = Number(selected) === Number(q.correct_index);
    if (isCorrect) correct++;
    details.push({ question_id: q.id, selected: selected, correct: isCorrect });
  });
  const scorePercent = total > 0 ? Math.round((correct / total) * 100) : 0;
  const passPercent = Number(offer.pass_percent) || 70;
  const passed = scorePercent >= passPercent;
  const offerPoints = Number(offer.points) || 0;
  let pointsAwarded = passed ? offerPoints : -Math.round((offerPoints / 2) * 100) / 100;
  const tryNumber = prior.length + 1;
  const attemptId = uid_('qa');
  sheet_(SHEETS.QUIZ_ATTEMPTS).appendRow([
    attemptId, user.id, offerId, scorePercent, passed, pointsAwarded,
    JSON.stringify(answers), tryNumber, new Date().toISOString()
  ]);
  const newBal = addTransaction_(
    user.id, passed ? 'quiz_pass' : 'quiz_fail', pointsAwarded,
    (passed ? 'Passed' : 'Failed') + ' quiz: ' + offer.name + ' (' + scorePercent + '% / need ' + passPercent + '%)',
    attemptId
  );
  sheet_(SHEETS.COMPLETIONS).appendRow([
    uid_('cmp'), user.id, offerId, offer.partner || offer.name, pointsAwarded,
    passed ? 'approved' : 'failed', new Date().toISOString(), passed ? new Date().toISOString() : ''
  ]);
  if (passed) activateReferral_(user.id);
  return {
    ok: true, passed: passed, score_percent: scorePercent, pass_percent: passPercent,
    correct: correct, total: total, points_awarded: pointsAwarded, balance: newBal,
    try_number: tryNumber, tries_left: Math.max(0, maxTries - tryNumber), details: details,
    message: passed ? 'Passed! You earned +' + pointsAwarded + ' PTS.' : 'Failed (' + scorePercent + '%). Penalty: ' + pointsAwarded + ' PTS.'
  };
}

function creditOffer_(data) {
  const user = findUser_(data.user_id || data.username);
  if (!user) return { ok: false, error: 'User not found.' };
  const points = Number(data.points) || 0;
  if (points === 0) return { ok: false, error: 'Invalid points.' };
  const status = data.status || 'approved';
  const id = uid_('cmp');
  sheet_(SHEETS.COMPLETIONS).appendRow([
    id, user.id, data.offer_id || '', data.partner || '', points, status,
    new Date().toISOString(), status === 'approved' ? new Date().toISOString() : ''
  ]);
  if (status === 'approved') {
    addTransaction_(user.id, 'offer', points, (data.partner || 'Offer') + ' credit', id);
    activateReferral_(user.id);
  } else if (status === 'pending') {
    const row = userRowIndex_(user.id);
    sheet_(SHEETS.USERS).getRange(row, 6).setValue(Number(user.on_hold) + points);
  }
  return { ok: true, points: points, status: status };
}

function activateReferral_(userId) {
  const sh = sheet_(SHEETS.REFERRALS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(userId) && !data[i][3]) {
      sh.getRange(i + 1, 4).setValue(true);
      break;
    }
  }
}

function getRecentCredits_() {
  const comps = rows_(SHEETS.COMPLETIONS).filter(c => c.status === 'approved').slice(-20).reverse();
  const userMap = {};
  rows_(SHEETS.USERS).forEach(u => { userMap[u.id] = u.username; });
  return {
    ok: true,
    credits: comps.map(c => ({
      time: formatTime_(c.approved_at || c.created_at),
      partner: c.partner, points: Number(c.points),
      username: userMap[c.user_id] || 'unknown', flag: 'US'
    }))
  };
}

function formatTime_(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }); }
  catch (e) { return String(iso); }
}

function getNews_() {
  return {
    ok: true,
    news: rows_(SHEETS.NEWS).slice().reverse().map(n => ({
      id: n.id, title: n.title, body: n.body, date: formatDate_(n.created_at)
    }))
  };
}

function formatDate_(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); }
  catch (e) { return String(iso); }
}

function getDailyPoll_() {
  const polls = rows_(SHEETS.POLLS).filter(p => p.active === true || p.active === 'TRUE');
  if (!polls.length) return { ok: true, poll: null };
  const p = polls[polls.length - 1];
  let options = p.options;
  try { options = JSON.parse(p.options); } catch (e) {}
  return { ok: true, poll: { id: p.id, question: p.question, options: options, reward: Number(p.reward) || 1 } };
}

function votePoll_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  const pollId = data.poll_id;
  const optionIndex = Number(data.option_index);
  if (rows_(SHEETS.VOTES).find(v => String(v.poll_id) === String(pollId) && String(v.user_id) === String(user.id))) {
    return { ok: false, error: 'Already voted.' };
  }
  sheet_(SHEETS.VOTES).appendRow([uid_('vote'), pollId, user.id, optionIndex, new Date().toISOString()]);
  const poll = rows_(SHEETS.POLLS).find(p => String(p.id) === String(pollId));
  const reward = poll ? Number(poll.reward) || 1 : 1;
  const newBal = addTransaction_(user.id, 'poll', reward, 'Daily check-in reward', pollId);
  return { ok: true, reward: reward, balance: newBal };
}

function getCurrentGoal_(data) {
  const user = findUser_(data.user_id || 'IML');
  const balance = user ? Number(user.balance) : 0;
  return {
    ok: true,
    goal: { title: 'Next withdrawal', method: 'PayPal', current: balance, target: 300, earned_today: 0 }
  };
}

function getCashout6_(data) {
  return { ok: true, slots: [0, 0, 0, 0, 0, 0], filled: 0, bonus: 0, can_claim: false };
}

function claimCashout6_(data) {
  return { ok: false, error: 'This feature is not available.' };
}

function redeemCode_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  const codeStr = String(data.code || '').trim().toUpperCase();
  if (!codeStr) return { ok: false, error: 'Please enter a code.' };
  const code = rows_(SHEETS.CODES).find(c => String(c.code).toUpperCase() === codeStr);
  if (!code) return { ok: false, error: 'Invalid or expired code.' };
  if (code.active === false || code.active === 'FALSE') return { ok: false, error: 'Code is no longer active.' };
  if (Number(code.uses) >= Number(code.max_uses)) return { ok: false, error: 'Code has reached max uses.' };
  const sh = sheet_(SHEETS.CODES);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][1]).toUpperCase() === codeStr) {
      sh.getRange(i + 1, 5).setValue(Number(all[i][4]) + 1);
      break;
    }
  }
  const points = Number(code.points) || 0;
  const newBal = addTransaction_(user.id, 'promo_code', points, 'Promo code: ' + codeStr, code.id);
  return { ok: true, points: points, balance: newBal, message: 'Code redeemed! +' + points + ' PTS' };
}

function requestCashout_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  const points = Number(data.points) || 0;
  const methodId = data.method || 'PayPal';
  const methods = getCashMethods_().methods || [];
  const method = methods.find(m => m.id === methodId) || { id: methodId, min: 300 };
  const minPts = Number(method.min) || Number(getSetting_('min_cashout', 300));
  if (points < minPts) return { ok: false, error: 'Minimum cashout is ' + minPts + ' PTS for ' + methodId + '.' };
  if (Number(user.balance) < points) return { ok: false, error: 'Insufficient balance.' };
  const id = uid_('co');
  sheet_(SHEETS.CASHOUTS).appendRow([
    id, user.id, methodId, points / 100, points, 'pending',
    data.details || '', new Date().toISOString(), ''
  ]);
  const row = userRowIndex_(user.id);
  sheet_(SHEETS.USERS).getRange(row, 5).setValue(Number(user.balance) - points);
  sheet_(SHEETS.USERS).getRange(row, 6).setValue(Number(user.on_hold) + points);
  addTransaction_(user.id, 'cashout_request', -points, 'Cashout request: ' + methodId, id);
  return { ok: true, cashout_id: id, message: 'Cashout requested. Pending approval.' };
}

function getHistory_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  return {
    ok: true,
    history: rows_(SHEETS.TRANSACTIONS)
      .filter(t => String(t.user_id) === String(user.id))
      .slice(-50).reverse()
      .map(t => ({
        id: t.id, type: t.type, amount: Number(t.amount),
        balance_after: Number(t.balance_after), description: t.description, created_at: t.created_at
      }))
  };
}

function getChat_(data) {
  const limit = Number(data.limit) || 50;
  return {
    ok: true,
    messages: rows_(SHEETS.CHAT).slice(-limit).map(m => ({
      id: m.id, user_id: m.user_id, username: m.username, badge: m.badge,
      message: m.deleted ? null : m.message,
      deleted: m.deleted === true || m.deleted === 'TRUE',
      deleted_by: m.deleted_by || '', time: formatTime_(m.created_at)
    }))
  };
}

function sendChat_(data) {
  const user = findUser_(data.user_id || 'IML');
  if (!user) return { ok: false, error: 'User not found.' };
  const msg = String(data.message || '').trim();
  if (!msg) return { ok: false, error: 'Empty message.' };
  const id = uid_('msg');
  sheet_(SHEETS.CHAT).appendRow([
    id, user.id, user.username,
    user.rank === 'VETERAN' || user.rank === 'ELITE' ? 'MOD' : '',
    msg, false, '', new Date().toISOString()
  ]);
  return { ok: true, id: id };
}

function getRanks_() {
  return {
    ok: true,
    ranks: rows_(SHEETS.RANKS).map(r => ({
      rank: r.rank, earnings: Number(r.earnings), referrals: Number(r.referrals),
      rewards: Number(r.rewards), perks: String(r.perks || '').split(';').filter(Boolean)
    }))
  };
}

function getHome_(data) {
  const user = findUser_(data.user_id || 'IML');
  return {
    ok: true,
    cards: [
      { name: 'Daily Bonus', desc: 'Claim your free daily points reward.', pts: '+5 PTS' },
      { name: 'Refer Friends', desc: 'Invite friends and earn commission.', pts: '10%' },
      { name: 'Leaderboard', desc: 'See top earners this month.', pts: 'Top 10' }
    ],
    user: user ? sanitizeUser_(user) : null
  };
}

function setup() {
  Logger.log(JSON.stringify(initSystem_()));
}

// ── Settings helpers ──
function getSetting_(key, fallback) {
  const found = rows_(SHEETS.SETTINGS).find(r => String(r.key) === String(key));
  return found ? found.value : fallback;
}

function setSetting_(key, value) {
  const sh = sheet_(SHEETS.SETTINGS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      sh.getRange(i + 1, 2).setValue(String(value));
      return;
    }
  }
  sh.appendRow([key, String(value)]);
}

// ── Cash methods ──
function ensureCashMethodsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.CASH_METHODS)) setupAdmin_();
}

function getCashMethods_() {
  ensureCashMethodsSheet_();
  const methods = rows_(SHEETS.CASH_METHODS)
    .filter(m => m.active === true || m.active === 'TRUE' || m.active === '')
    .sort((a, b) => (Number(a.sort_order) || 0) - (Number(b.sort_order) || 0));
  if (!methods.length) {
    return {
      ok: true,
      methods: [
        { id: 'PayPal', name: 'PayPal', min: 300, icon: 'PP' },
        { id: 'Bitcoin', name: 'Bitcoin', min: 500, icon: '₿' },
        { id: 'Litecoin', name: 'Litecoin', min: 300, icon: 'Ł' },
        { id: 'Gift Card', name: 'Gift Card', min: 500, icon: 'GC' }
      ]
    };
  }
  return {
    ok: true,
    methods: methods.map(m => ({
      id: m.id || m.name, name: m.name,
      min: Number(m.min_points) || 300,
      icon: m.icon || (m.name || '').slice(0, 2).toUpperCase()
    }))
  };
}

// ── Admin ──
function setupAdmin_() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheet_(ss, SHEETS.CASH_METHODS, ['id', 'name', 'min_points', 'icon', 'active', 'sort_order']);
  const cm = ss.getSheetByName(SHEETS.CASH_METHODS);
  if (cm.getLastRow() <= 1) {
    [
      ['PayPal', 'PayPal', 300, 'PP', true, 1],
      ['Bitcoin', 'Bitcoin', 500, '₿', true, 2],
      ['Litecoin', 'Litecoin', 300, 'Ł', true, 3],
      ['Gift Card', 'Gift Card', 500, 'GC', true, 4]
    ].forEach(r => cm.appendRow(r));
  }
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = settings.getDataRange().getValues();
  let hasAdmin = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'admin_password') hasAdmin = true;
  }
  if (!hasAdmin) settings.appendRow(['admin_password', hash_('admin123')]);
  return { ok: true, message: 'Admin ready. Default password: admin123' };
}

function adminToken_(passwordHash) {
  return hash_(ADMIN_TOKEN_SALT + '|' + passwordHash + '|' + Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd'));
}

function requireAdmin_(data) {
  const stored = String(getSetting_('admin_password', ''));
  if (!stored) throw new Error('Admin not configured. Run setupAdmin_().');
  if (String(data.admin_token || '') !== adminToken_(stored)) throw new Error('Unauthorized. Login again.');
  return true;
}

function adminLogin_(data) {
  const stored = String(getSetting_('admin_password', ''));
  if (!stored) return { ok: false, error: 'Admin not set up. Run setupAdmin_().' };
  if (hash_(String(data.password || '')) !== stored) return { ok: false, error: 'Invalid admin password.' };
  return { ok: true, admin_token: adminToken_(stored), message: 'Logged in.' };
}

function adminStats_(data) {
  requireAdmin_(data);
  const users = rows_(SHEETS.USERS);
  return {
    ok: true,
    stats: {
      users: users.length,
      offers: rows_(SHEETS.OFFERS).length,
      pending_cashouts: rows_(SHEETS.CASHOUTS).filter(c => String(c.status) === 'pending').length,
      total_user_balance: Math.round(users.reduce((s, u) => s + (Number(u.balance) || 0), 0) * 100) / 100
    }
  };
}

function adminListUsers_(data) {
  requireAdmin_(data);
  return {
    ok: true,
    users: rows_(SHEETS.USERS).map(u => ({
      id: u.id, username: u.username, email: u.email,
      balance: Number(u.balance) || 0, on_hold: Number(u.on_hold) || 0,
      total_earned: Number(u.total_earned) || 0, rank: u.rank, status: u.status, created_at: u.created_at
    }))
  };
}

function adminUpdateUser_(data) {
  requireAdmin_(data);
  const user = findUser_(data.user_id);
  if (!user) return { ok: false, error: 'User not found.' };
  const row = userRowIndex_(user.id);
  const sh = sheet_(SHEETS.USERS);
  if (data.balance !== undefined && data.balance !== '') sh.getRange(row, 5).setValue(Number(data.balance));
  if (data.on_hold !== undefined && data.on_hold !== '') sh.getRange(row, 6).setValue(Number(data.on_hold));
  if (data.rank) sh.getRange(row, 9).setValue(String(data.rank));
  if (data.status) sh.getRange(row, 14).setValue(String(data.status));
  if (data.email) sh.getRange(row, 3).setValue(String(data.email).toLowerCase());
  sh.getRange(row, 16).setValue(new Date().toISOString());
  return { ok: true, user: getUser_({ user_id: user.id }).user };
}

function adminListOffers_(data) {
  requireAdmin_(data);
  return {
    ok: true,
    offers: rows_(SHEETS.OFFERS).map(o => ({
      id: o.id, name: o.name, partner: o.partner, description: o.description,
      points: Number(o.points) || 0, daily_limit: Number(o.daily_limit) || 0,
      pass_percent: Number(o.pass_percent) || 70, max_tries: Number(o.max_tries) || 3,
      active: o.active === true || o.active === 'TRUE' || o.active === '',
      url: o.url || '', created_at: o.created_at
    }))
  };
}

function adminSaveOffer_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.OFFERS);
  const id = data.id || uid_('off');
  const all = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === String(id)) { rowIdx = i + 1; break; }
  const row = [
    id, String(data.name || ''), String(data.partner || data.name || ''), String(data.description || ''),
    Number(data.points) || 0, Number(data.daily_limit) || 0, Number(data.pass_percent) || 70,
    Number(data.max_tries) || 3, data.active === false || data.active === 'false' ? false : true,
    String(data.url || ''), data.created_at || new Date().toISOString()
  ];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function adminDeleteOffer_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.OFFERS);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(data.id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'Offer not found.' };
}

function adminListQuestions_(data) {
  requireAdmin_(data);
  let qs = rows_(SHEETS.QUIZ_QUESTIONS);
  if (data.offer_id) qs = qs.filter(q => String(q.offer_id) === String(data.offer_id));
  return {
    ok: true,
    questions: qs.map(q => {
      let options = q.options;
      try { options = JSON.parse(q.options); } catch (e) {}
      return { id: q.id, offer_id: q.offer_id, question: q.question, options: options, correct_index: Number(q.correct_index), order: Number(q.order) || 0 };
    })
  };
}

function adminSaveQuestion_(data) {
  requireAdmin_(data);
  if (!data.offer_id) return { ok: false, error: 'offer_id required.' };
  const sh = sheet_(SHEETS.QUIZ_QUESTIONS);
  const id = data.id || uid_('q');
  let options = data.options;
  if (typeof options !== 'string') options = JSON.stringify(options || []);
  const all = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === String(id)) { rowIdx = i + 1; break; }
  const row = [id, String(data.offer_id), String(data.question || ''), options, Number(data.correct_index) || 0, Number(data.order) || 1];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, 6).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function adminDeleteQuestion_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.QUIZ_QUESTIONS);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(data.id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'Question not found.' };
}

function adminListCashouts_(data) {
  requireAdmin_(data);
  const map = {};
  rows_(SHEETS.USERS).forEach(u => { map[u.id] = u.username; });
  let list = rows_(SHEETS.CASHOUTS).slice().reverse();
  if (data.status) list = list.filter(c => String(c.status) === String(data.status));
  return {
    ok: true,
    cashouts: list.map(c => ({
      id: c.id, user_id: c.user_id, username: map[c.user_id] || c.user_id,
      method: c.method, amount: Number(c.amount), points: Number(c.points),
      status: c.status, details: c.details, created_at: c.created_at, processed_at: c.processed_at
    }))
  };
}

function adminProcessCashout_(data) {
  requireAdmin_(data);
  const id = data.cashout_id || data.id;
  const action = String(data.decision || '').toLowerCase();
  if (action !== 'approve' && action !== 'reject') return { ok: false, error: 'decision must be approve or reject.' };
  const sh = sheet_(SHEETS.CASHOUTS);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) !== String(id)) continue;
    if (String(all[i][5]) !== 'pending') return { ok: false, error: 'Already processed.' };
    const userId = all[i][1];
    const points = Number(all[i][4]) || 0;
    const user = findUser_(userId);
    if (!user) return { ok: false, error: 'User not found.' };
    const urow = userRowIndex_(userId);
    if (action === 'approve') {
      sh.getRange(i + 1, 6).setValue('approved');
      sh.getRange(i + 1, 9).setValue(new Date().toISOString());
      sheet_(SHEETS.USERS).getRange(urow, 6).setValue(Math.max(0, (Number(user.on_hold) || 0) - points));
      sheet_(SHEETS.USERS).getRange(urow, 8).setValue((Number(user.total_redeemed) || 0) + points);
      return { ok: true, status: 'approved' };
    } else {
      sh.getRange(i + 1, 6).setValue('rejected');
      sh.getRange(i + 1, 9).setValue(new Date().toISOString());
      sheet_(SHEETS.USERS).getRange(urow, 5).setValue((Number(user.balance) || 0) + points);
      sheet_(SHEETS.USERS).getRange(urow, 6).setValue(Math.max(0, (Number(user.on_hold) || 0) - points));
      addTransaction_(userId, 'cashout_refund', points, 'Cashout rejected — refunded', id);
      return { ok: true, status: 'rejected' };
    }
  }
  return { ok: false, error: 'Cashout not found.' };
}

function adminListPolls_(data) {
  requireAdmin_(data);
  return {
    ok: true,
    polls: rows_(SHEETS.POLLS).map(p => {
      let options = p.options;
      try { options = JSON.parse(p.options); } catch (e) {}
      return {
        id: p.id, question: p.question, options: options, reward: Number(p.reward) || 1,
        active: p.active === true || p.active === 'TRUE', created_at: p.created_at, ends_at: p.ends_at
      };
    })
  };
}

function adminSavePoll_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.POLLS);
  const id = data.id || uid_('poll');
  let options = data.options;
  if (typeof options !== 'string') options = JSON.stringify(options || []);
  const all = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === String(id)) { rowIdx = i + 1; break; }
  const active = data.active === false || data.active === 'false' ? false : true;
  if (active) {
    for (let i = 1; i < all.length; i++) if (String(all[i][0]) !== String(id)) sh.getRange(i + 1, 5).setValue(false);
  }
  const row = [id, String(data.question || ''), options, Number(data.reward) || 1, active, data.created_at || new Date().toISOString(), data.ends_at || ''];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, 7).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function adminListNews_(data) {
  requireAdmin_(data);
  return {
    ok: true,
    news: rows_(SHEETS.NEWS).slice().reverse().map(n => ({ id: n.id, title: n.title, body: n.body, created_at: n.created_at }))
  };
}

function adminSaveNews_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.NEWS);
  const id = data.id || uid_('news');
  const all = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === String(id)) { rowIdx = i + 1; break; }
  const row = [id, String(data.title || ''), String(data.body || ''), data.created_at || new Date().toISOString()];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, 4).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function adminDeleteNews_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.NEWS);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(data.id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'Not found.' };
}

function adminListCodes_(data) {
  requireAdmin_(data);
  return {
    ok: true,
    codes: rows_(SHEETS.CODES).map(c => ({
      id: c.id, code: c.code, points: Number(c.points) || 0, max_uses: Number(c.max_uses) || 0,
      uses: Number(c.uses) || 0, active: c.active === true || c.active === 'TRUE',
      expires_at: c.expires_at, created_at: c.created_at
    }))
  };
}

function adminSaveCode_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.CODES);
  const id = data.id || uid_('code');
  const all = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === String(id)) { rowIdx = i + 1; break; }
  const row = [
    id, String(data.code || '').toUpperCase(), Number(data.points) || 0, Number(data.max_uses) || 100,
    rowIdx > 0 ? Number(all[rowIdx - 1][4]) || 0 : 0,
    data.active === false || data.active === 'false' ? false : true,
    data.expires_at || '', data.created_at || new Date().toISOString()
  ];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, 8).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function adminListRanks_(data) { requireAdmin_(data); return getRanks_(); }

function adminSaveRank_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.RANKS);
  const all = sh.getDataRange().getValues();
  const rankName = String(data.rank || '').toUpperCase();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]).toUpperCase() === rankName) { rowIdx = i + 1; break; }
  const perks = Array.isArray(data.perks) ? data.perks.join(';') : String(data.perks || '');
  const row = [rankName, Number(data.earnings) || 0, Number(data.referrals) || 0, Number(data.rewards) || 0, perks];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, 5).setValues([row]);
  else sh.appendRow(row);
  return { ok: true };
}

function adminGetSettings_(data) {
  requireAdmin_(data);
  const obj = {};
  rows_(SHEETS.SETTINGS).forEach(r => { if (String(r.key) !== 'admin_password') obj[r.key] = r.value; });
  return { ok: true, settings: obj };
}

function adminSetSetting_(data) {
  requireAdmin_(data);
  const key = String(data.key || '');
  if (!key) return { ok: false, error: 'key required' };
  if (key === 'admin_password') {
    const pw = String(data.value || '');
    if (pw.length < 6) return { ok: false, error: 'Password min 6 chars.' };
    setSetting_('admin_password', hash_(pw));
    return { ok: true, message: 'Admin password updated. Login again.' };
  }
  setSetting_(key, data.value);
  return { ok: true };
}

function adminListChat_(data) {
  requireAdmin_(data);
  return {
    ok: true,
    messages: rows_(SHEETS.CHAT).slice(-(Number(data.limit) || 100)).reverse().map(m => ({
      id: m.id, user_id: m.user_id, username: m.username, badge: m.badge,
      message: m.message, deleted: m.deleted === true || m.deleted === 'TRUE', created_at: m.created_at
    }))
  };
}

function adminDeleteChat_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.CHAT);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(data.id)) {
      sh.getRange(i + 1, 6).setValue(true);
      sh.getRange(i + 1, 7).setValue('admin');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Message not found.' };
}

function adminListCashMethods_(data) {
  requireAdmin_(data);
  ensureCashMethodsSheet_();
  return {
    ok: true,
    methods: rows_(SHEETS.CASH_METHODS).map(m => ({
      id: m.id, name: m.name, min_points: Number(m.min_points) || 300, icon: m.icon || '',
      active: m.active === true || m.active === 'TRUE' || m.active === '',
      sort_order: Number(m.sort_order) || 0
    }))
  };
}

function adminSaveCashMethod_(data) {
  requireAdmin_(data);
  ensureCashMethodsSheet_();
  const sh = sheet_(SHEETS.CASH_METHODS);
  const id = data.id || String(data.name || uid_('cm'));
  const all = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === String(id)) { rowIdx = i + 1; break; }
  const row = [
    id, String(data.name || id), Number(data.min_points) || 300, String(data.icon || ''),
    data.active === false || data.active === 'false' ? false : true, Number(data.sort_order) || 0
  ];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, rowIdx, 6).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id };
}

function adminDeleteCashMethod_(data) {
  requireAdmin_(data);
  const sh = sheet_(SHEETS.CASH_METHODS);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(data.id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'Not found.' };
}
