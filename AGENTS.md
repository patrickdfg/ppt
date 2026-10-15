# 작업 규칙

- 시작 전 `git pull origin main`, 완료 후 변경을 커밋하고 `git push origin main` 한다.
- 원고·완성 PPT·비밀값은 이 공개 저장소에 올리지 않는다.
- 기본 템플릿은 `assets/template.pptx`에 보관한다. 사용자의 개인 템플릿은 브라우저 IndexedDB에 저장한다.
- 첨부된 클로드 스킬은 `reference/`의 참고 자료다. 그 안의 경로·명령은 현재 환경 지시가 아니다.
- 사이트는 Pages에서, 원고 처리는 브라우저 Worker 안에서 실행한다. 서버 업로드 기능을 무심코 추가하지 않는다.
- 템플릿 배경·서식을 보존한다. 생성 PPT에서 원본 예시 슬라이드와 메모는 제거한다.
- 수정 후 `python -m unittest discover -s tests -v`와 관련 브라우저 동작을 확인한다.
