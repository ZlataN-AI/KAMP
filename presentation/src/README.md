# RIVET-WM 세미나 덱 생성 스크립트

`../RIVET-WM_seminar.pptx`를 만드는 pptxgenjs 스크립트입니다 (GAT 세미나 스타일).

- `build.js` : 슬라이드 구성·도형 다이어그램·차트 정의
- `notes.json` : 원본 발표자료의 발표자 노트
- `media/` : 원본 발표자료에서 가져온 결과 그림과 수식 이미지

```bash
npm install pptxgenjs@3.12.0
node build.js ../RIVET-WM_seminar.pptx
```

`build.js`는 pptx 스킬의 `apply_theme.js`(`/mnt/skills/public/pptx/scripts/`)로 테마 색을 기록합니다.
