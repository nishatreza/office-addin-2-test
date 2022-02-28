import { Component } from "@angular/core";

/* global Word */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";

  async run() {
    return Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
      paragraph.font.color = "blue";

      await context.sync();
    });
  }

  async addText() {
    return Word.run(async (context) => {
      const inputVal = context.document.getElementById("myInput").value;
      // const paragraph = context.document.body.insertParagraph(inputVal, Word.InsertLocation.end);
      // paragraph.font.color = "black";
      var docBody = context.document.body;
      docBody.insertParagraph(inputVal);
      await context.sync();
    });
  }

  async changeFontBold() {
    return Word.run(async (context) => {
      var selection = context.document.getSelection();

      selection.font.bold = true;
      await context.sync();
    });
  }

  async changeFontItalic() {
    return Word.run(async (context) => {
      var selection = context.document.getSelection();

      selection.font.italic = true;
      await context.sync();
    });
  }
}
