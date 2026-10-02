// Verifica se o assunto contém um código de projeto no formato [KD-123]
var PADRAO = /\[KD-\d+\]/i;

function liberar(event) {
  try { event.completed({ allowEvent: true }); } catch (e) {}
}

function onMessageSendHandler(event) {
  try {
    Office.context.mailbox.item.subject.getAsync(function (result) {
      try {
        // Se não conseguir ler o assunto, não bloqueia o envio
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          liberar(event);
          return;
        }

        var assunto = result.value || "";

        if (PADRAO.test(assunto)) {
          event.completed({ allowEvent: true });
        } else {
          event.completed({
            allowEvent: false,
            errorMessage:
              "O assunto não tem o ID do projeto. Adicione no formato [KD-123] para facilitar a busca depois."
          });
        }
      } catch (e) {
        liberar(event);
      }
    });
  } catch (e) {
    liberar(event);
  }
}

Office.onReady(function () {});
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
