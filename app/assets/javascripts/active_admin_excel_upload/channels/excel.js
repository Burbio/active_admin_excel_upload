App.room = App.cable.subscriptions.create("ActiveAdminExcelUpload::ExcelChannel", {
  connected: function() {
    // Called when the subscription is ready for use on the server
    return console.log('connected');
  },
  disconnected: function() {},
  // Called when the subscription has been terminated by the server
  received: function(data) {
    var model, modelInMessage, para, searchParams, t;
    // Called when there's incoming data on the websocket for this channel
    searchParams = new URLSearchParams(window.location.search);
    model = searchParams.get("model");
    modelInMessage = data.match("\\[(\\w+\)]")[1];
    if (model === modelInMessage) {
      para = document.createElement("P");
      t = document.createTextNode(data.match("\\[(\\w+\)](.*)")[2]);
      para.appendChild(t);
      para.style.borderBottom = "1px solid lightgrey";
      para.style.margin = "0px";
      para.style.padding = "10px";
      return document.getElementById("show_status").appendChild(para);
    }
  },
  speak: function() {
    return this.perform('speak');
  }
});
