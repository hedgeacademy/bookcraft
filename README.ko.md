# Bookcraft

Bookcraft는 긴 중국어 도서 프로젝트를 기획, 집필, 검토, 패키징, 내보내기까지 지원하는 Codex Skill입니다. 상업적 표지 브리프와 Word/DOCX 내보내기 workflow도 포함합니다.

## 저자 및 연락처

Created by Xinyi Chen, founder of HEDGE Global.

중국어 소개: 陈歆怡，海聚海外 CEO。

Contact: `chenxinyi_g`

## 주요 기능

- 책의 포지셔닝, 독자 약속, 제목, 부제, 전체 목차 설계.
- `book.toml`, `src/SUMMARY.md`, 관리 파일, 장 파일을 포함한 mdBook 스타일 프로젝트 생성.
- 자료, 인용, 검증 필요 항목, 참고문헌 상태 관리.
- 장별 집필, 확장, 검토, 수정과 연속성 유지.
- 출판 구조 정리: 서문, 본문, 참고문헌, 선택 후속 자료.
- 상업적 표지 패키징: 판매 훅, 시각 상징, 제목 계층, 서가 테스트, 썸네일 테스트.
- `scripts/md2word.py`를 통한 `manuscript.md`의 Word/DOCX 내보내기.

## 설치

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/hedgeacademy/bookcraft ~/.codex/skills/bookcraft
```

새 Codex 스레드에서 다음처럼 사용합니다.

```text
Use $bookcraft to plan a book about [topic].
```

## 예시 프롬프트

```text
Use $bookcraft to create a book project from this outline.
```

```text
Use $bookcraft to review this chapter for structure, evidence, and continuity.
```

```text
Use $bookcraft to build a commercial cover brief for this manuscript.
```

```text
Use $bookcraft to export this manuscript to Word/DOCX.
```

## 표지 방법론

표지는 단순히 보기 좋은 이미지가 아니라 상업적으로 즉시 이해되는 신호여야 합니다. 좋은 표지는 몇 초 안에 책의 주제, 독자, 약속, 선택 이유를 전달합니다.

참고:

- `references/09-cover-design.ko.md`
- `references/09-cover-design.md`

## 내보내기 요구 사항

DOCX 내보내기에는 Python, `python-docx`, `pandoc`, 완성된 `manuscript.md`, 유효한 `book.toml`, 실제 참고문헌 섹션이 필요합니다.

## 라이선스

MIT
