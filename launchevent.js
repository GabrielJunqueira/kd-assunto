// Verifica se o assunto contém um código de projeto no formato [KD-123]
const PADRAO = /\[KD-\d+\]/i;

function onMessageSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, (result) => {
    const ev = result.asyncContext;

    // Se não conseguir ler o assunto, não bloqueia o envio
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      ev.completed({ allowEvent: true });
      return;
    }

    const assunto = result.value || "";

    if (PADRAO.test(assunto)) {
      ev.completed({ allowEvent: true });
    } else {
      ev.completed({
        allowEvent: false,
        errorMessage:
          "O assunto não tem o ID do projeto. Adicione no formato [KD-123] para facilitar a busca depois."
      });
    }
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
