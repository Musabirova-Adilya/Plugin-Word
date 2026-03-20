Office.onReady(() => {
  const button = document.getElementById("formatButton");
  const status = document.getElementById("status");
  const select = document.getElementById("styleSelect");

  status.textContent = "НОВАЯ ВЕРСИЯ ПЛАГИНА ЗАГРУЖЕНА";

  button.onclick = async () => {
    try {
      status.textContent = "Ищу таблицу...";

      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tables.items.length === 0) {
          status.textContent = "В документе нет таблиц.";
          return;
        }

        const table = tables.items[0];
        const style = select.value;

        table.styleBuiltIn = "TableGrid";

        const rows = table.rows;
        rows.load("items");
        await context.sync();

        for (let i = 0; i < rows.items.length; i++) {
          const row = rows.items[i];
          const cells = row.cells;
          cells.load("items");
          await context.sync();

          for (let j = 0; j < cells.items.length; j++) {
            const cell = cells.items[j];
            const range = cell.body.getRange();

            range.font.color = "black";

            if (style === "simple") {
              range.font.name = "Calibri";
              range.font.size = 11;
              range.font.bold = false;
            } else {
              range.font.name = "Times New Roman";
              range.font.size = 12;
              range.font.bold = (i === 0);
            }

            if (style === "espd") {
              const paragraphs = cell.body.paragraphs;
              paragraphs.load("items");
              await context.sync();

              for (let k = 0; k < paragraphs.items.length; k++) {
                paragraphs.items[k].alignment = "Centered";
              }
            }
          }
        }

        await context.sync();
        status.textContent = "Таблица отформатирована";
      });

    } catch (error) {
      status.textContent = "Ошибка: " + error.message;
      console.error(error);
    }
  };
});