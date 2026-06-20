# Small Group Guide Generator / 소그룹교안 생성기

A tool for **churches and missionary organizations** — upload a sermon file and AI automatically generates a structured small group / cell group / home group guide.

**교회와 선교단체**를 위한 도구입니다. 설교 파일(PDF · DOCX · HWPX · TXT)을 업로드하면 AI가 자동으로 소그룹(마을예배·셀·목장·구역) 교안을 작성해 줍니다.

> **The UI automatically displays in Korean on Korean Windows and English on English Windows.**
> 한글 윈도우에서는 한국어, 영문 윈도우에서는 영어 UI로 자동 표시됩니다.

## Features / 주요 기능

- **Bilingual UI** — Korean on Korean Windows, English on English Windows / 한글 윈도우 한국어 · 영문 윈도우 영어 자동 전환
- **Multiple AI engines** — Claude CLI · Codex CLI · Anthropic API · Gemini API · ChatGPT API
- **Sermon file input** — PDF · DOCX · HWPX · TXT / 설교 파일 입력
- **Output formats** — TXT · DOCX · MD / 출력 형식 선택
- **Rich preview** — visual formatting in DOCX/MD mode / 서식 미리보기
- **API Key storage** — per-engine key saving / 백엔드별 Key 개별 저장
- **Progress timer** — shows elapsed time during generation / 진행 타이머
- **CLI install helper** — detect & install Claude CLI · Codex CLI / CLI 설치 도우미

## How to Run / 실행 방법

### EXE (no Python needed / Python 불필요)
```
smallgroup-guide-generator.exe  ←  double-click to run / 더블클릭으로 실행
```
[Download latest release](https://github.com/Brightinyou/smallgroup-guide-generator/releases)

### Python
```bash
pip install -r requirements.txt
python app.py
```

## Build EXE / EXE 빌드

```bat
build.bat
```
Outputs `dist/smallgroup-guide-generator.exe`

## AI Engine Requirements / AI 엔진별 요구사항

| Engine | Requirement |
|--------|-------------|
| Claude CLI | [Claude Code](https://claude.ai/code) installed |
| Codex CLI | Node.js + `npm install -g @openai/codex` |
| Anthropic API | Anthropic API Key |
| Gemini API | Google AI API Key |
| ChatGPT API | OpenAI API Key |

## Config File / 설정 파일

`~/.village_worship_cfg.json` — stores API Keys, save folder, last-used engine

## Guide Structure / 출력 교안 구조

**Korean (한국어)**
1. 마을예배 속으로 (아이스브레이크 2문항)
2. 말씀 속으로
3. 삶 속으로 (설교 요약 + 나눔 질문 2세트)
4. 합심기도 (기도제목 4항목)

**English**
1. Getting Started (2 icebreaker questions)
2. Into the Word
3. Applying the Word (sermon summary + 2 discussion questions)
4. Prayer Together (4 prayer points)
