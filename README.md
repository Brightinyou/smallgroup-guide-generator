# Small Group Guide Generator / 소그룹교안 생성기

A tool for **churches and missionary organizations** — upload a sermon file and AI automatically generates a structured small group / cell group / home group guide.

**교회와 선교단체**를 위한 도구입니다. 설교 파일(PDF · DOCX · HWPX · TXT)을 업로드하면 AI가 자동으로 소그룹(마을예배·셀·목장·구역) 교안을 작성해 줍니다.

> **The UI automatically displays in Korean on Korean Windows and English on English Windows.**
> 한글 윈도우에서는 한국어, 영문 윈도우에서는 영어 UI로 자동 표시됩니다.

---

## 💡 No Extra API Charges / 추가 과금 없이 사용하기

**If you already subscribe to Claude or ChatGPT, you can use this app for free by installing the CLI — no API key, no extra billing.**

클로드(Claude) 또는 챗GPT 구독자라면, CLI를 설치하면 **추가 API 과금 없이** 구독 한도 안에서 사용할 수 있습니다.

| Subscription / 구독 | Method / 방법 | Engine to select / 앱에서 선택 |
|---|---|---|
| Claude Pro / Max | Install Claude Code CLI | `Claude CLI` |
| ChatGPT Plus / Pro | Install Codex CLI | `Codex CLI` |
| No subscription | Use API Key (pay-per-use) | Anthropic / Gemini / ChatGPT API |

---

## 🟣 Claude Code CLI — for Claude Pro/Max subscribers

**Claude Pro 또는 Max 구독자 전용 — API Key 불필요**

### Prerequisites / 사전 준비

- **Node.js 18+** — [nodejs.org](https://nodejs.org) 에서 LTS 버전 다운로드

### Installation / 설치

```bash
npm install -g @anthropic-ai/claude-code
```

> Windows에서 권한 오류 시: PowerShell을 **관리자 권한**으로 열고 실행하세요.

### First-time login / 최초 로그인

```bash
claude
```

실행하면 브라우저가 열리며 Anthropic 계정 로그인 화면이 나옵니다.
Claude.ai에서 사용하는 계정(Pro/Max 구독 계정)으로 로그인하면 인증이 완료됩니다.

로그인 후 터미널로 돌아와 `Ctrl+C`로 종료해도 인증 정보는 유지됩니다.

### Using in the app / 앱에서 사용

1. 앱 실행 → **AI 엔진 선택**에서 `Claude CLI` 선택
2. 연결됨 메시지(`Connected: ...claude.exe`)가 초록색으로 표시되면 정상
3. 설교 파일 선택 → **교안 생성하기** 클릭

> **권장 모델**: `claude-sonnet-4-6` (기본값) — Pro/Max 구독 한도 내 사용

---

## 🟢 Codex CLI — for ChatGPT Plus/Pro subscribers

**ChatGPT Plus 또는 Pro 구독자 전용 — API Key 불필요**

Codex CLI는 OpenAI가 배포한 오픈소스 CLI로, 구독 계정으로 로그인하면 ChatGPT 구독 한도 안에서 GPT 모델을 사용할 수 있습니다.

### Prerequisites / 사전 준비

- **Node.js 18+** — [nodejs.org](https://nodejs.org) 에서 LTS 버전 다운로드

### Installation / 설치

```bash
npm install -g @openai/codex
```

### First-time login / 최초 로그인

```bash
codex
```

실행하면 브라우저가 열리며 OpenAI 계정 로그인 화면이 나옵니다.
ChatGPT.com에서 사용하는 계정(Plus/Pro 구독 계정)으로 로그인하면 인증이 완료됩니다.

> **ChatGPT Plus 구독자**: `gpt-5.5` 모델 사용 가능
> **ChatGPT Pro 구독자**: `o4-mini`, `gpt-4o` 등 추가 모델 사용 가능

### Using in the app / 앱에서 사용

1. 앱 실행 → **AI 엔진 선택**에서 `Codex CLI` 선택
2. 연결됨 메시지(`Connected: ...codex.cmd`)가 초록색으로 표시되면 정상
3. **모델** 선택:
   - `gpt-5.5` — ChatGPT Plus/Pro 구독자 (기본값, 추가 과금 없음)
   - `o4-mini` / `gpt-4o` — API 계정 또는 Pro 구독자
4. 설교 파일 선택 → **교안 생성하기** 클릭

### Install helper / 설치 도우미

앱 오른쪽 상단의 **설치 도우미** 버튼을 클릭하면 Node.js / Claude CLI / Codex CLI 설치 상태를 확인하고 npm으로 자동 설치할 수 있습니다.

---

## 🔑 Using API Keys / API Key 사용 (구독 없는 경우)

구독이 없는 경우 각 서비스의 API Key를 직접 발급받아 사용합니다.

| Engine | API Key 발급 |
|--------|-------------|
| Anthropic API | [console.anthropic.com](https://console.anthropic.com) |
| Gemini API | [aistudio.google.com](https://aistudio.google.com) |
| ChatGPT API | [platform.openai.com](https://platform.openai.com) |

앱에서 엔진 선택 후 API Key 입력란에 붙여넣고 **저장** 버튼을 클릭하세요.

---

## 📥 Download & Run / 다운로드 및 실행

### EXE (no Python needed / Python 불필요)

[**⬇ Download latest release**](https://github.com/Brightinyou/smallgroup-guide-generator/releases)

`smallgroup-guide-generator.exe` 를 다운로드 후 더블클릭하면 바로 실행됩니다.
Python 설치 불필요. Windows 10/11 지원.

### Python (개발자용)

```bash
pip install -r requirements.txt
python app.py
```

### Build EXE / EXE 빌드

```bat
build.bat
```

---

## 📋 Guide Structure / 교안 구조

**Korean (한국어)**
1. 마을예배 속으로 — 아이스브레이크 2문항
2. 말씀 속으로
3. 삶 속으로 — 설교 요약 + 나눔 질문 2세트
4. 합심기도 — 기도제목 4항목

**English**
1. Getting Started — 2 icebreaker questions
2. Into the Word
3. Applying the Word — sermon summary + 2 discussion questions
4. Prayer Together — 4 prayer points

**Output formats / 출력 형식**: TXT · DOCX · Markdown

---

## ⚙️ Config / 설정 파일

`~/.village_worship_cfg.json` — API Keys, save folder, last-used engine are saved here automatically.
API Key, 저장 폴더, 마지막 사용 엔진이 자동 저장됩니다.
