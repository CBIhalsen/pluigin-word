// 插入 OMML 数学公式的函数
async function insertOMML() {
    try {
        await Word.run(async (context) => {
            // 插入 OMML 数学公式
            const omml = `
                <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
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
            `;

            // 将 OMML 插入到选定位置
            const range = context.document.getSelection();
            range.insertOoxml(omml, Word.InsertLocation.replace);

            await context.sync();
            console.log("数学公式已插入！");
        });
    } catch (error) {
        console.error("Error:", error);
    }
}
