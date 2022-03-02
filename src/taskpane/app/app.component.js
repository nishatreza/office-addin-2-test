import { Component } from "@angular/core";

/* global Word */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  // eslint-disable-next-line prettier/prettier
  clickMessage = '';

  // async run() {
  //   return Word.run(async (context) => {
  //     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

  //     // change the paragraph color to blue.
  //     paragraph.font.color = "blue";

  //     await context.sync();
  //   });
  // }

  async onClickShowText(input) {
    return Word.run(async (context) => {
      this.clickMessage = input;

      const paragraph = context.document.body.insertParagraph(this.clickMessage, Word.InsertLocation.end);
      paragraph.font.color = "blue";

      await context.sync();
    });
  }

  async onKeyTypeText(event) {
    return Word.run(async (context) => {
      const inputValue = event.target.value;

      const paragraph = context.document.body.insertParagraph(inputValue, Word.InsertLocation.end);


      paragraph.font.color = "blue";

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
