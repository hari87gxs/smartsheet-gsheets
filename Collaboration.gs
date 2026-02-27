// ============================================================================
// Collaboration.gs — Row-level comments, @mentions, sharing helpers
// ============================================================================

var COMMENTS_META_KEY = 'PS_COMMENTS';   // developer metadata key on each row

// ── Comments ──────────────────────────────────────────────────────────────────
/**
 * Adds a comment to a row. Detects @mentions and sends notification emails.
 *
 * @param {number} rowIndex  — 1-based
 * @param {string} commentText
 */
function addComment(rowIndex, commentText) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var user  = Session.getActiveUser().getEmail();
  var ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");

  var existing = getComments(rowIndex);
  existing.push({ user: user, text: commentText, ts: ts });
  _setRowMeta(sheet, rowIndex, COMMENTS_META_KEY, JSON.stringify(existing));

  // Handle @mentions
  var mentions = commentText.match(/@([\w.]+)/g) || [];
  mentions.forEach(function(m) {
    var email = m.slice(1);
    if (!email.includes('@')) email += '@' + (user.split('@')[1] || 'example.com');
    _sendMentionEmail(email, user, commentText, sheet, rowIndex);
  });

  logActivity(sheet, rowIndex, 'COMMENT', 'Comment by ' + user.split('@')[0] + ': ' + commentText.slice(0,60));
  return existing;
}

/**
 * Returns all comments for a row.
 */
function getComments(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var val   = _getRowMeta(sheet, rowIndex, COMMENTS_META_KEY);
  return val ? JSON.parse(val) : [];
}

/**
 * Deletes a comment by index.
 */
function deleteComment(rowIndex, commentIndex) {
  var comments = getComments(rowIndex);
  var sheet    = SpreadsheetApp.getActiveSheet();
  var user     = Session.getActiveUser().getEmail();
  if (comments[commentIndex] && (comments[commentIndex].user === user ||
      sheet.getOwner().getEmail() === user)) {
    comments.splice(commentIndex, 1);
    _setRowMeta(sheet, rowIndex, COMMENTS_META_KEY, JSON.stringify(comments));
  }
  return comments;
}

// ── Row metadata helpers ──────────────────────────────────────────────────────
function _setRowMeta(sheet, rowIndex, key, value) {
  var range = sheet.getRange(rowIndex, 1);
  // Remove old
  var existing = range.getDeveloperMetadata();
  existing.forEach(function(m){ if (m.getKey() === key) m.remove(); });
  // Add new
  range.addDeveloperMetadata(key, value, SpreadsheetApp.DeveloperMetadataVisibility.PROJECT);
}

function _getRowMeta(sheet, rowIndex, key) {
  var metas = sheet.getRange(rowIndex, 1).getDeveloperMetadata();
  for (var i=0;i<metas.length;i++){
    if (metas[i].getKey() === key) return metas[i].getValue();
  }
  return null;
}

// ── Comment sidebar HTML ──────────────────────────────────────────────────────
/**
 * Opens the comment panel for the currently selected row.
 */
function openCommentPanel() {
  var rowIndex = SpreadsheetApp.getActiveSheet().getActiveRange().getRow();
  if (rowIndex <= 1) {
    SpreadsheetApp.getUi().alert('Select a task row first.');
    return;
  }
  var html = HtmlService.createHtmlOutput(buildCommentHtml(rowIndex))
    .setTitle('Row Comments')
    .setWidth(400)
    .setHeight(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildCommentHtml(rowIndex) {
  var comments = getComments(rowIndex);
  var sheet    = SpreadsheetApp.getActiveSheet();
  var taskName = '';
  try {
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var taskCol = headers.indexOf('Task Name');
    if (taskCol >= 0) taskName = sheet.getRange(rowIndex, taskCol+1).getValue();
  } catch(e){}

  var commentsHtml = comments.length ? comments.map(function(c, i) {
    var initials = (c.user||'?').split('@')[0].slice(0,2).toUpperCase();
    var relTime  = _relativeTime(new Date(c.ts));
    return '<div class="comment">' +
      '<div class="c-header"><div class="avatar">'+esc(initials)+'</div>' +
      '<div class="c-meta"><b>'+esc((c.user||'').split('@')[0])+'</b>' +
      '<span class="c-time">'+relTime+'</span></div>' +
      '</div><div class="c-body">'+_formatMentions(esc(c.text))+'</div></div>';
  }).join('') : '<div class="empty">No comments yet. Be the first! 💬</div>';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{box-sizing:border-box;margin:0;padding:0;}' +
    'body{font-family:-apple-system,sans-serif;background:#1a1a2e;color:#e8eaed;height:100vh;display:flex;flex-direction:column;}' +
    '#header{padding:12px 16px;background:#16213e;border-bottom:1px solid #0f3460;}' +
    '#header h3{font-size:13px;color:#4fc3f7;}' +
    '#header p{font-size:11px;color:#9aa0a6;margin-top:2px;}' +
    '#comments{flex:1;overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:10px;}' +
    '.comment{background:#0d1b2a;border-radius:10px;padding:12px;}' +
    '.c-header{display:flex;align-items:center;gap:8px;margin-bottom:8px;}' +
    '.avatar{width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#1a73e8,#4fc3f7);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;}' +
    '.c-meta{display:flex;flex-direction:column;}' +
    '.c-meta b{font-size:12px;}' +
    '.c-time{font-size:10px;color:#9aa0a6;}' +
    '.c-body{font-size:12px;line-height:1.5;}' +
    '.mention{color:#4fc3f7;font-weight:600;}' +
    '.empty{color:#555;font-size:12px;text-align:center;padding:20px;}' +
    '#compose{padding:12px;background:#16213e;border-top:1px solid #0f3460;}' +
    'textarea{width:100%;background:#0d1b2a;border:1px solid #0f3460;color:#e8eaed;padding:8px;border-radius:8px;font-size:12px;resize:none;outline:none;font-family:inherit;}' +
    'textarea:focus{border-color:#1a73e8;}' +
    '#compose-row{display:flex;gap:8px;margin-top:8px;}' +
    'button{padding:6px 14px;border-radius:6px;border:none;cursor:pointer;font-size:12px;}' +
    '.btn-primary{background:#1a73e8;color:#fff;flex:1;}' +
    '.btn-primary:hover{background:#1557c0;}' +
    '.hint{font-size:10px;color:#555;margin-top:4px;}' +
    '</style></head><body>' +
    '<div id="header"><h3>💬 Comments — Row ' + rowIndex + '</h3>' +
    (taskName ? '<p>'+esc(String(taskName))+'</p>' : '') + '</div>' +
    '<div id="comments">' + commentsHtml + '</div>' +
    '<div id="compose">' +
    '<textarea id="msg" rows="3" placeholder="Add a comment… use @name to mention"></textarea>' +
    '<div class="compose-row"><button class="btn-primary" onclick="send()">Send</button></div>' +
    '<div class="hint">Tip: @mention someone to notify them by email</div>' +
    '</div>' +
    '<script>' +
    'function send(){var t=document.getElementById("msg").value.trim();if(!t)return;' +
    'google.script.run.withSuccessHandler(function(){location.reload();}).addComment(' + rowIndex + ',t);}' +
    'document.getElementById("msg").addEventListener("keydown",function(e){if(e.ctrlKey&&e.key==="Enter")send();});' +
    '</scr' + 'ipt></body></html>';
}

// ── @mention email ────────────────────────────────────────────────────────────
function _sendMentionEmail(to, from, commentText, sheet, rowIndex) {
  try {
    if (!_isEmail(to)) return;
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var url    = ss.getUrl();
    GmailApp.sendEmail(to,
      'You were mentioned in ' + ss.getName(),
      from + ' mentioned you in row ' + rowIndex + ' of "' + sheet.getName() + '":\n\n' +
      '"' + commentText + '"\n\nOpen sheet: ' + url
    );
    logActivity(sheet, rowIndex, 'COMMENT_MENTION', 'Mentioned ' + to);
  } catch(err) {
    Logger.log('[Mention Email Error] ' + err.message);
  }
}

// ── Sharing helper ────────────────────────────────────────────────────────────
/**
 * Shares the current spreadsheet with an email address.
 */
function shareWithUser(email, role) {
  role = role || 'commenter';       // 'viewer' | 'commenter' | 'editor'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.addEditor(email);
    logActivity(null, 0, 'SHARE', 'Shared with ' + email + ' as ' + role);
    return true;
  } catch(err) {
    Logger.log('[Share Error] ' + err.message);
    return false;
  }
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function _relativeTime(date) {
  var diff = Math.round((new Date() - date) / 1000);
  if (diff < 60)   return 'just now';
  if (diff < 3600) return Math.floor(diff/60) + 'm ago';
  if (diff < 86400)return Math.floor(diff/3600) + 'h ago';
  return Math.floor(diff/86400) + 'd ago';
}

function _formatMentions(text) {
  return text.replace(/@([\w.]+)/g, '<span class="mention">@$1</span>');
}

function _isEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }
function esc(s)      { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
