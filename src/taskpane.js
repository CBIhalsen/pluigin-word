// taskpane.js

// Office 初始化
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // 绑定按钮事件
    document.getElementById("insertTextButton").onclick = insertText;
    document.getElementById("insertOMMLButton").onclick = insertOMML;

    console.log("插件已初始化");
  }
});

// 插入普通文本
async function insertText() {
  try {
    await Word.run(async (context) => {
      // 获取当前选区
      const range = context.document.getSelection();
      // 在选区中插入文本
      range.insertText("Hello World!", Word.InsertLocation.replace);

      await context.sync();
      console.log("文本已插入！");
    });
  } catch (error) {
    console.error("插入文本时出错:", error);
  }
}

// 插入 OMML 数学公式
async function insertOMML() {
  try {
    await Word.run(async (context) => {
      const omml = `
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:r>
    <m:oMath>
      <m:r><m:t>E</m:t></m:r>
      <m:r><m:t>=</m:t></m:r>
      <m:r><m:t>m</m:t></m:r>
      <m:sSup>
        <m:e><m:r><m:t>c</m:t></m:r></m:e>
        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
      </m:sSup>
    </m:oMath>
  </w:r>
</w:p>
`;

      const range = context.document.getSelection();
      range.insertOoxml(omml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.error("插入公式时出错:", error);
    console.error("详细调试信息:", error.debugInfo);
  }
}