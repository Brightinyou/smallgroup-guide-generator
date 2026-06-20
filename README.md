# 소그룹교안 생성기

설교 파일(PDF · DOCX · HWPX · TXT)을 업로드하면 AI가 자동으로 소그룹(마을예배·셀·목장·구역) 교안을 작성해주는 도구입니다.

## 주요 기능

- **다양한 AI 엔진 지원**: Claude CLI · Codex CLI · Anthropic API · Gemini API · ChatGPT API
- **설교 파일 입력**: PDF · DOCX · HWPX · TXT
- **출력 형식 선택**: TXT · DOCX · MD
- **서식 미리보기**: DOCX/MD 모드에서 시각적 서식 적용
- **API Key 저장**: 백엔드별 Key 개별 저장
- **진행 타이머**: 교안 생성 중 경과 시간 표시
- **CLI 설치 도우미**: Claude CLI · Codex CLI 설치 상태 확인 및 설치

## 실행 방법

### EXE (Python 불필요)
```
소그룹교안생성기.exe  ←  더블클릭으로 실행
```
[최신 릴리즈 다운로드](https://github.com/Brightinyou/smallgroup-guide-generator/releases)

### Python으로 직접 실행
```bash
pip install -r requirements.txt
python 교안생성기.py
```

## EXE 빌드

```bat
빌드EXE.bat
```
`dist/소그룹교안생성기.exe` 가 생성됩니다.

## AI 엔진별 요구사항

| 엔진 | 필요 사항 |
|------|-----------|
| Claude CLI | [Claude Code](https://claude.ai/code) 설치 |
| Codex CLI | Node.js + `npm install -g @openai/codex` |
| Anthropic API | Anthropic API Key |
| Gemini API | Google AI API Key |
| ChatGPT API | OpenAI API Key |

## 설정 파일

`~/.village_worship_cfg.json` 에 저장 (API Key · 저장 폴더 · 마지막 사용 엔진 등)

## 출력 교안 구조

1. 마을예배 속으로 (아이스브레이크 2문항)
2. 말씀 속으로
3. 삶 속으로 (설교 요약 + 나눔 질문 2세트)
4. 합심기도 (기도제목 4항목)
