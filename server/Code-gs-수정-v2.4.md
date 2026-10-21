# Code.gs v2.3 → v2.4 수정 안내 (받을 학생 학번·이름 저장)

촬영 페이지 v2.0이 보내는 `recipient`(대표 학생 학번·이름)를 운영대장에 저장합니다.
**`CONFIG`(관리자 코드·이메일 등)는 건드리지 않습니다.** 아래 다섯 곳만 바꿉니다.

- 순서는 상관없지만, 다섯 곳을 모두 바꾼 뒤 저장하고 배포해야 합니다.
- 기존 운영대장 시트는 그대로 씁니다. 첫 접수 때 맨 오른쪽(M열)에 `recipient` 제목이 자동으로 생깁니다.
- 촬영 페이지를 먼저 올려도 문제없습니다. 옛 서버는 `recipient`를 무시하고 그대로 접수합니다. 이 경우 학번·이름만 저장되지 않습니다.

---

## ① `HEADERS` 맨 끝에 한 줄 추가

찾기:
```js
  'printedAt',
  'doneAt'
]);
```
바꾸기:
```js
  'printedAt',
  'doneAt',
  'recipient'
]);
```

## ② `validateUpload_` — 반환값 바로 위에 정리 코드 추가

찾기:
```js
  return {
    jobId,
    boothId,
    characterId,
```
바꾸기:
```js
  // 사진 배부용 대표 학생 학번·이름 (선택). 제어 문자 제거, 20자 제한.
  // 시트 수식으로 해석되지 않도록 =,+,-,@ 로 시작하면 앞에 작은따옴표를 붙입니다.
  let recipient = String(p.recipient || '')
    .replace(/[\u0000-\u001f\u007f<>]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 20);
  if (/^[=+\-@]/.test(recipient)) recipient = "'" + recipient;

  return {
    recipient,
    jobId,
    boothId,
    characterId,
```

## ③ `doPost` — `appendRow` 맨 끝에 한 칸 추가

찾기:
```js
          CONFIG.STATUS_WAITING,
          '',
          ''
        ]);
```
바꾸기:
```js
          CONFIG.STATUS_WAITING,
          '',
          '',
          validated.recipient
        ]);
```

## ④ `ensureHeader_` — 함수 전체를 아래로 교체

```js
function ensureHeader_(sheet, canWrite) {
  if (sheet.getLastRow() === 0) {
    if (!canWrite) {
      throw publicError_('SERVER_NOT_READY', '운영대장 헤더가 없습니다.');
    }

    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS.slice()]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    return;
  }

  const current = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getDisplayValues()[0];

  if (HEADERS.every((h, i) => current[i] === h)) return;

  // v2.3 시트: 앞 12칸이 같고 마지막 칸(recipient)만 비어 있으면 제목 한 칸만 채웁니다.
  const last = HEADERS.length - 1;
  const legacy = HEADERS.slice(0, last).every((h, i) => current[i] === h) && !current[last];
  if (legacy) {
    sheet.getRange(1, last + 1).setValue(HEADERS[last]).setFontWeight('bold');
    return;
  }

  throw new Error(
    `시트 헤더가 v2.4 형식과 다릅니다. "${CONFIG.SHEET_NAME}" 탭을 확인해 주세요.`
  );
}
```

## ⑤ `getPrintItem` — 반환값에 한 줄 추가

찾기:
```js
  return {
    jobId: row.jobId,
    boothId: row.boothId,
```
바꾸기:
```js
  return {
    jobId: row.jobId,
    recipient: row.recipient || '',
    boothId: row.boothId,
```

---

## 저장과 배포

1. 저장(Ctrl+S)
2. **배포 → 배포 관리** → 지금 쓰는 배포의 연필 → 버전 **새 버전** → 배포. 주소는 그대로입니다.
3. 관리자 화면(`Admin.html`)도 v2.6으로 바꿨다면 같은 배포 하나로 함께 반영됩니다.

## 확인

1. 촬영 페이지에서 한 장 찍고 "사진 받을 학생"에 `20315 홍길동`처럼 입력해 접수합니다.
2. 운영대장 시트 M열 `recipient`에 그 값이 들어갔는지 봅니다.
3. 관리자 화면 '받을 학생' 칸에 보이는지, 검색 칸에 `홍길동`이나 `20315`를 쳐서 찾아지는지 봅니다.

## 개인정보

- 입력 칸 아래에 "사진을 나눠 줄 때만 쓰고, 행사가 끝나면 사진과 함께 삭제해요. 사진에는 들어가지 않아요."라고 안내합니다.
- 관리자 화면의 **행사 데이터 일괄 삭제**를 하면 운영대장의 행이 모두 지워지므로 학번·이름도 함께 지워집니다.
- 학번·이름은 사진 파일 이름에도, 인쇄되는 사진에도 들어가지 않습니다.
