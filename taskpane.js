Office.onReady(() => {
  const input = document.getElementById("prefix");
  const msg = document.getElementById("msg");
  input.value = Office.context.roamingSettings.get("prefix") || "";

  document.getElementById("save").onclick = () => {
    if (!/^\d*$/.test(input.value)) {
      msg.textContent = "数字のみ入力してください";
      return;
    }
    Office.context.roamingSettings.set("prefix", input.value);
    Office.context.roamingSettings.saveAsync(() => (msg.textContent = "保存しました"));
  };

  document.getElementById("run").onclick = run;
});

function yymmdd(d = new Date()) {
  const p = (n) => String(n).padStart(2, "0");
  return p(d.getFullYear() % 100) + p(d.getMonth() + 1) + p(d.getDate());
}

async function run() {
  const prefix = Office.context.roamingSettings.get("prefix") || "";
  await Excel.run(async (ctx) => {
    const cell = ctx.workbook.getActiveCell();
    cell.numberFormat = [["@"]];
    cell.values = [[prefix + yymmdd()]];
    await ctx.sync();
  });
}
