let email = {
  toEmail: "",
  subject: "",
  body: "",
};

function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync("text", { asyncContext: event }, getBodyCallback);
  //   Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, getSubjectCallback);
  //   Office.context.mailbox.item.to.getAsync({ asyncContext: event }, getToCallback);
}

function getToCallback(asyncResult) {
  const event = asyncResult.asyncContext;
  let to = "";

  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    to = asyncResult.value;
  } else {
    const message = "Failed to get to text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  console.log("To:" + to[0].emailAddress);
  email.toEmail = to[0].emailAddress;

  const userAction = async () => {
    const response = await fetch("http://localhost:5016/api/EmailStorage/SaveEmail", {
      method: "POST",
      body: {
        email: email,
      },
      headers: {
        accept: "*/*",
        "Access-Control-Allow-Origin": "https://localhost:3000/",
        "Content-Type": "application/json",
      },
    });
    const myJson = await response.json(); //extract JSON from the http response
    // do something with myJson
  };

  userAction();

  event.completed({ allowEvent: true });
}

function getSubjectCallback(asyncResult) {
  const event = asyncResult.asyncContext;
  let subject = "";

  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    subject = asyncResult.value;
  } else {
    const message = "Failed to get subject text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  console.log("Subject: " + subject);
  email.subject = subject;

  Office.context.mailbox.item.to.getAsync({ asyncContext: event }, getToCallback);
}

function getBodyCallback(asyncResult) {
  const event = asyncResult.asyncContext;
  let body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    const message = "Failed to get body text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  console.log("Body: " + body);
  email.body = body;

  Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, getSubjectCallback);
}

if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
