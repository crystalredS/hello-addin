// Office.js 初期化
Office.onReady(() => {
  console.log("Add-in ready");
});

/**
 * DocumentOpened イベントで呼ばれる関数
 * マニフェストの FunctionName="onDocumentOpened" と一致させる
 */
function onDocumentOpened(event) {
  console.log("DocumentOpened event fired");

  // 文書の状態をチェック（例：先頭のテキストを読む）
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text,
    (result) => {
      console.log("Document content:", result.value);

      // メッセージを Office の通知バーに表示
      Office.context.ui.displayDialogAsync(
        "https://crystalreds.github.io/hello-addin/message.html",
        { height: 30, width: 30 },
        (asyncResult) => {
          console.log("Dialog opened");
        }
      );

      // イベント完了
      event.completed();
    }
  );
}
