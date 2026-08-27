// ============================================================
// 스마트 발주 시스템 - 웹앱 URL 설정
// ------------------------------------------------------------
// Apps Script를 "새 배포"가 아니라 "기존 배포 수정(관리 → 배포 수정 → 새 버전)"으로
// 업데이트하면 아래 URL은 그대로 유지됩니다. 그래도 주소가 바뀌었다면
// 이 파일에서 해당 줄만 새 URL로 바꾸고 저장(Commit)하면 됩니다.
// index.html 본문 코드는 전혀 건드릴 필요 없습니다.
// ============================================================
window.APP_CONFIG = {
  // 예정량/확정량 시트(부대별정보/품목정보/공휴일캐시/데이터/DB_Master가 있는 스프레드시트)
  HISTORY_WEBAPP_URL: "https://script.google.com/macros/s/AKfycbw46wO5Yww9RavXDXaOcVCeysUkoEqTHkXGJCH5ozfsnRwJEIm9S4AXnrJPPP9QeXQQqQ/exec",

  // 재고수불대장 시트(국탕류/식육가공품/자사양념육/납품기한임박/납품기한경과 등이 있는 스프레드시트)
  LEDGER_WEBAPP_URL: "https://script.google.com/macros/s/AKfycbyYuCI_jY0S6yAFwxjZqmYjGmG7LXgalbedUqorFuL7GtadInkVl7NN9q7JjK7Vuec/exec"
};
