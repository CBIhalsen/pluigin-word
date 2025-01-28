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
      // 这里演示一个最简单的 E = mc^2 示例
      // 注意使用 <m:oMathPara> 作为顶层标签，以便 Word 能正确解析
      const omml = `
<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath>
    <m:r>
      <m:t>E</m:t>
    </m:r>
    <m:r>
      <m:t>=</m:t>
    </m:r>
    <m:r>
      <m:t>m</m:t>
    </m:r>
    <m:sSup>
      <m:e>
        <m:r>
          <m:t>c</m:t>
        </m:r>
      </m:e>
      <m:sup>
        <m:r>
          <m:t>2</m:t>
        </m:r>
      </m:sup>
    </m:sSup>
  </m:oMath>
</m:oMathPara>
`;

      // 获取当前选区
      const range = context.document.getSelection();
      // 插入 OMML
      range.insertOoxml(omml, Word.InsertLocation.replace);

      await context.sync();
      console.log("数学公式已插入！");
    });
  } catch (error) {
    console.error("插入公式时出错:", error);
  }
}