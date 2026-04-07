// Google Chat notifications (Spaces) via Incoming Webhooks.
// This keeps email notifications intact and adds Chat as an additional channel.

var BIBLIOTECA_URL = 'https://sites.google.com/telesalud.gob.sv/bibliotecadivisioneinnovacion/p%C3%A1gina-principal';

function sendChatNotification_(action, articleObj) {
  try {
    var webhooks = getChatWebhooks_();
    if (!webhooks || !webhooks.length) return;

    var base = getBasePublicUrl_();
    var url = base ? (base + '#/articulo/' + articleObj.slug) : ('#/articulo/' + articleObj.slug);
    var message = buildChatMessage_(action, articleObj, url);

    for (var i = 0; i < webhooks.length; i++) {
      sendToChatWebhook_(webhooks[i], message);
    }
  } catch (e) {
    try {
      logAudit_('chat', (articleObj && articleObj.articleId) || '', 'ERROR', getActiveUserEmail_(), {
        message: e && e.message ? e.message : String(e)
      });
    } catch (_) {}
  }
}

function buildChatMessage_(action, articleObj, url) {
  var title = (articleObj && articleObj.title) || 'Sin título';
  var author = (articleObj && (articleObj.updatedBy || articleObj.createdBy)) || '';
  if (!author) {
    try { author = getActiveUserEmail_(); } catch (_) {}
  }
  if (!author) author = 'usuario@local';

  var actionLabel = action === 'create' ? 'Creación de artículo' : action === 'update' ? 'Actualización de artículo' : 'Eliminación de artículo';
  var date = toIsoString_(now_()).slice(0, 19).replace('T', ' ');

  // Descripción: se puede enviar HTML limitado (solo etiquetas que Chat admite)
  var descHtml = (articleObj && articleObj.descriptionHtmlSanitized) ? String(articleObj.descriptionHtmlSanitized) : '';
  var descForChat = htmlToChatFormat_(descHtml, 240);

  var card = {
    cards: [{
      header: {
        title: 'Biblioteca DoctorSV',
        subtitle: 'Notificación institucional',
        imageUrl: 'https://i.pinimg.com/originals/57/9b/90/579b90f5e64d7631fc81e55ba716ba9f.png',
        imageStyle: 'AVATAR'
      },
      sections: [{
        widgets: [
          {
            textParagraph: {
              text: '<b>' + escapeHtml_(title) + '</b>'
            }
          },
          {
            keyValue: {
              topLabel: 'Acción',
              content: escapeHtml_(actionLabel),
              icon: 'DESCRIPTION'
            }
          },
          {
            keyValue: {
              topLabel: 'Autor',
              content: escapeHtml_(author),
              icon: 'PERSON'
            }
          },
          {
            keyValue: {
              topLabel: 'Fecha',
              content: escapeHtml_(date),
              icon: 'CLOCK'
            }
          }
        ]
      }]
    }]
  };

  if (articleObj && articleObj.driveMimeType) {
    card.cards[0].sections[0].widgets.push({
      keyValue: {
        topLabel: 'Categoría',
        content: escapeHtml_(String(articleObj.driveMimeType)),
        icon: 'BOOKMARK'
      }
    });
  }

  if (descForChat) {
    card.cards[0].sections[0].widgets.push({
      textParagraph: { text: descForChat }
    });
  }

  if (action !== 'delete' && url) {
    card.cards[0].sections.push({
      widgets: [{
        buttons: [{
          textButton: {
            text: 'Ver artículo',
            onClick: { openLink: { url: url } }
          }
        }]
      }]
    });
  }

  // En todos los mensajes: enlace a la Biblioteca (página principal).
  card.cards[0].sections.push({
    widgets: [{
      buttons: [{
        textButton: {
          text: 'Ir a la Biblioteca',
          onClick: { openLink: { url: BIBLIOTECA_URL } }
        }
      }]
    }]
  });

  return {
    text: 'Notificación de artículo: ' + actionLabel + ' - ' + title,
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
    throw new Error('Chat webhook failed: HTTP ' + code + ' - ' + response.getContentText());
  }
}

function getChatWebhooks_() {
  var json = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOKS');
  if (!json) return [];
  try {
    var list = JSON.parse(json);
    if (!Array.isArray(list)) return [];
    return list.filter(function(item) { return item && item.url; });
  } catch (_) {
    return [];
  }
}

// Public helper: call once to register a webhook.
function setupChatWebhook(webhookUrl, spaceName) {
  if (!webhookUrl || !String(webhookUrl).trim()) return 'Webhook requerido';
  var list = getChatWebhooks_();
  var url = String(webhookUrl).trim();
  var name = String(spaceName || 'Espacio principal').trim();
  var exists = false;
  for (var i = 0; i < list.length; i++) {
    if (list[i].url === url) {
      list[i].name = name;
      exists = true;
      break;
    }
  }
  if (!exists) list.push({ name: name, url: url });
  PropertiesService.getScriptProperties().setProperty('CHAT_WEBHOOKS', JSON.stringify(list));
  return 'Webhook guardado: ' + name;
}

// Public helper: list configured webhooks (masked).
function listChatWebhooks() {
  var list = getChatWebhooks_();
  return list.map(function(item) {
    var masked = item.url ? (item.url.slice(0, 32) + '...') : '';
    return { name: item.name || 'Espacio', url: masked };
  });
}

// Public helper: notificación de prueba (mismo formato e institucional que las reales).
function testChatNotification() {
  var author = '';
  try { author = getActiveUserEmail_(); } catch (_) {}
  if (!author) author = 'usuario@local';

  var base = getBasePublicUrl_();
  var linkUrl = base ? (base + '#/articulo/notificacion-de-prueba') : '#/articulo/notificacion-de-prueba';

  var testArticle = {
    articleId: 'test-' + uuid_(),
    title: 'Notificación de prueba - Biblioteca DoctorSV',
    slug: 'notificacion-de-prueba',
    descriptionHtmlSanitized: '<p><b>Mensaje de prueba</b> del canal institucional.</p><p>Formato: <i>cursiva</i>, <b>negrita</b>. Enlace de ejemplo: <a href="' + linkUrl + '">Ver artículo</a>. No requiere acción.</p>',
    driveMimeType: 'Comunicación institucional',
    createdBy: author,
    updatedBy: author
  };

  sendChatNotification_('update', testArticle);
  return 'Se envió la notificación de prueba a Google Chat con el formato institucional configurado.';
}

// Public helper: remove a webhook.
function removeChatWebhook(webhookUrl) {
  var url = String(webhookUrl || '').trim();
  if (!url) return 'Webhook requerido';
  var list = getChatWebhooks_();
  var next = list.filter(function(item) { return item.url !== url; });
  PropertiesService.getScriptProperties().setProperty('CHAT_WEBHOOKS', JSON.stringify(next));
  return next.length < list.length ? 'Webhook eliminado' : 'Webhook no encontrado';
}

function stripHtml_(value) {
  return String(value || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function truncateText_(value, maxLen) {
  var text = String(value || '').trim();
  if (!text) return '';
  if (!maxLen || text.length <= maxLen) return text;
  return text.slice(0, maxLen - 3) + '...';
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Convierte HTML a formato que Google Chat admite en textParagraph.
 * Solo se conservan: <b>, <i>, <u>, <s>, <br>, <font color="...">, <a href="...">.
 * El resto de etiquetas se eliminan (el texto interno se mantiene).
 */
function htmlToChatFormat_(html, maxLen) {
  if (!html || !String(html).trim()) return '';
  var text = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<p[^>]*>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<div[^>]*>/gi, '\n')
    .replace(/&nbsp;/gi, ' ');
  // Quitar etiquetas no permitidas (dejar solo b, i, u, s, font, a)
  text = text.replace(/<(?!\/?(?:b|i|u|s|font|a)\b)[^>]*>/gi, '');
  text = text.replace(/\n{2,}/g, '\n').replace(/^\s+|\s+$/g, '').trim();
  if (maxLen && text.length > maxLen) {
    text = text.slice(0, maxLen - 3) + '...';
  }
  return text;
}
