// ============================================================
//  ChatService.gs — Biblioteca de Telemedicina
//  Notificaciones a Google Chat via Incoming Webhooks.
// ============================================================

function sendChatNotification_(action, articleObj) {
  try {
    var webhooks = getChatWebhooks_();
    if (!webhooks || !webhooks.length) return;

    var base    = getBasePublicUrl_();
    var url     = base ? (base + '#/articulo/' + articleObj.slug) : ('#/articulo/' + articleObj.slug);
    var message = buildChatMessage_(action, articleObj, url);

    for (var i = 0; i < webhooks.length; i++) {
      sendToChatWebhook_(webhooks[i], message);
    }
  } catch (e) {
    try {
      logAudit_('chat', (articleObj && articleObj.articleId) || '', 'ERROR',
        getActiveUserEmail_(), { message: e && e.message ? e.message : String(e) });
    } catch (_) {}
  }
}

function buildChatMessage_(action, articleObj, url) {
  var title  = (articleObj && articleObj.title) || 'Sin título';
  var author = (articleObj && (articleObj.updatedBy || articleObj.createdBy)) || '';
  if (!author) { try { author = getActiveUserEmail_(); } catch (_) {} }
  if (!author) author = 'sistema@telesalud.gob.sv';

  var actionLabel = action === 'create' ? 'Nuevo recurso publicado'
                  : action === 'update' ? 'Recurso actualizado'
                  : 'Recurso eliminado';

  var date        = toIsoString_(now_()).slice(0, 19).replace('T', ' ');
  var descHtml    = (articleObj && articleObj.descriptionHtmlSanitized)
                    ? String(articleObj.descriptionHtmlSanitized) : '';
  var descForChat = htmlToChatFormat_(descHtml, 240);

  var card = {
    cards: [{
      header: {
        title:    'Biblioteca Telemedicina',
        subtitle: 'Notificación institucional – División de Innovación',
        imageUrl: 'https://www.telesalud.gob.sv/wp-content/uploads/2020/09/logo-telesalud.png',
        imageStyle: 'AVATAR'
      },
      sections: [{
        widgets: [
          { textParagraph: { text: '<b>' + escapeHtml_(title) + '</b>' } },
          { keyValue: { topLabel: 'Acción',    content: escapeHtml_(actionLabel), icon: 'DESCRIPTION' } },
          { keyValue: { topLabel: 'Publicado por', content: escapeHtml_(author), icon: 'PERSON' } },
          { keyValue: { topLabel: 'Fecha',     content: escapeHtml_(date),       icon: 'CLOCK' } }
        ]
      }]
    }]
  };

  if (articleObj && articleObj.driveMimeType) {
    card.cards[0].sections[0].widgets.push({
      keyValue: { topLabel: 'Categoría', content: escapeHtml_(String(articleObj.driveMimeType)), icon: 'BOOKMARK' }
    });
  }

  if (descForChat) {
    card.cards[0].sections[0].widgets.push({ textParagraph: { text: descForChat } });
  }

  if (action !== 'delete' && url) {
    card.cards[0].sections.push({
      widgets: [{ buttons: [{ textButton: {
        text: 'Ver recurso',
        onClick: { openLink: { url: url } }
      }}]}]
    });
  }

  card.cards[0].sections.push({
    widgets: [{ buttons: [{ textButton: {
      text: 'Ir a la Biblioteca',
      onClick: { openLink: { url: getBasePublicUrl_() } }
    }}]}]
  });

  return {
    text: 'Biblioteca Telemedicina: ' + actionLabel + ' — ' + title,
    cards: card.cards
  };
}

function sendToChatWebhook_(webhook, message) {
  if (!webhook || !webhook.url) throw new Error('Webhook inválido');
  var options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(webhook.url, options);
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Chat webhook falló: HTTP ' + code + ' - ' + response.getContentText());
  }
}

function getChatWebhooks_() {
  var json = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOKS');
  if (!json) return [];
  try {
    var list = JSON.parse(json);
    if (!Array.isArray(list)) return [];
    return list.filter(function(item) { return item && item.url; });
  } catch (_) { return []; }
}

function setupChatWebhook(webhookUrl, spaceName) {
  if (!webhookUrl || !String(webhookUrl).trim()) return 'Webhook requerido';
  var list  = getChatWebhooks_();
  var url   = String(webhookUrl).trim();
  var name  = String(spaceName || 'Biblioteca Telemedicina').trim();
  var found = false;
  for (var i = 0; i < list.length; i++) {
    if (list[i].url === url) { list[i].name = name; found = true; break; }
  }
  if (!found) list.push({ name: name, url: url });
  PropertiesService.getScriptProperties().setProperty('CHAT_WEBHOOKS', JSON.stringify(list));
  return 'Webhook guardado: ' + name;
}

function listChatWebhooks() {
  return getChatWebhooks_().map(function(item) {
    return { name: item.name || 'Espacio', url: item.url ? (item.url.slice(0, 32) + '...') : '' };
  });
}

function removeChatWebhook(webhookUrl) {
  var url  = String(webhookUrl || '').trim();
  if (!url) return 'Webhook requerido';
  var list = getChatWebhooks_();
  var next = list.filter(function(item) { return item.url !== url; });
  PropertiesService.getScriptProperties().setProperty('CHAT_WEBHOOKS', JSON.stringify(next));
  return next.length < list.length ? 'Webhook eliminado' : 'No encontrado';
}

function testChatNotification() {
  var author = '';
  try { author = getActiveUserEmail_(); } catch (_) {}
  if (!author) author = 'sistema@telesalud.gob.sv';
  var linkUrl = getBasePublicUrl_() + '#/articulo/notificacion-de-prueba';
  var testArticle = {
    articleId:                'test-' + uuid_(),
    title:                    'Notificación de prueba — Biblioteca Telemedicina',
    slug:                     'notificacion-de-prueba',
    descriptionHtmlSanitized: '<p><b>Mensaje de prueba</b> del canal institucional de Telemedicina.</p>',
    driveMimeType:            'Comunicación institucional',
    createdBy:  author,
    updatedBy:  author
  };
  sendChatNotification_('update', testArticle);
  return 'Notificación de prueba enviada a Google Chat.';
}

// ----------------------------------------------------------------
// Helpers de texto
// ----------------------------------------------------------------
function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function htmlToChatFormat_(html, maxLen) {
  if (!html || !String(html).trim()) return '';
  var text = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n').replace(/<p[^>]*>/gi, '\n')
    .replace(/<\/div>/gi, '\n').replace(/<div[^>]*>/gi, '\n')
    .replace(/&nbsp;/gi, ' ');
  text = text.replace(/<(?!\/?(?:b|i|u|s|font|a)\b)[^>]*>/gi, '');
  text = text.replace(/\n{2,}/g, '\n').trim();
  if (maxLen && text.length > maxLen) text = text.slice(0, maxLen - 3) + '...';
  return text;
}