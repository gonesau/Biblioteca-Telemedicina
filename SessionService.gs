function getSessionInfo() {
  var email = getActiveUserEmail_();
  var cfg = getAccessConfig();
  var isAdmin = cfg.adminEmails && cfg.adminEmails.indexOf(email) !== -1;
  var isGuestEditor = cfg.guestEditorEmails && cfg.guestEditorEmails.indexOf(email) !== -1;
  var allowed = isAdmin || isGuestEditor || isEmailAllowed_(email);
  return { email: email, isAdmin: isAdmin, isGuestEditor: isGuestEditor, allowed: allowed, basePublicUrl: getBasePublicUrl_() };
}


