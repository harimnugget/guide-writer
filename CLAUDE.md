# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

개인 스터디/사이드 프로젝트 공간. 빌드 시스템·의존성 없는 순수 HTML/CSS/JS 프로젝트. 각 `.html` 파일을 브라우저에서 직접 열면 됨.

## 파일 구성

- `calculator.html` — 고양이 테마 계산기 (단일 파일, HTML+CSS+JS)
- `sidiz-presentation.html` — SIDIZ GX & GC PRO 제품 프레젠테이션 (키보드 내비게이션, Pretendard 폰트 CDN)
- `sidiz-presentation.pptx` — 위 프레젠테이션의 PPTX 버전

## calculator.html 구조

**계산기 상태 변수 (JS)**
- `current` — 현재 입력값 (문자열)
- `previous` — 이전 피연산자
- `operator` — 현재 연산자 (`+`, `-`, `*`, `/`)
- `shouldReset` — 다음 입력 시 current를 초기화할지 여부

**주요 함수**
- `inputDigit(digit)` / `inputDot()` — 숫자·소수점 입력
- `setOperator(op)` — 연산자 지정, 연속 연산 시 자동으로 `calculate(true)` 호출
- `calculate(chained?)` — 결과 계산; 0 나누기는 `'오류'` 표시
- `clearAll()` / `toggleSign()` / `percent()` — AC, +/-, % 기능

**인터랙션**
- 키보드 입력 지원 (숫자, 사칙연산, Enter/=, Escape, Backspace, `/` 기본동작 차단)
- 마우스 이동에 따라 계산기 3D 기울기 효과 적용

**버튼 스타일 클래스**
- `.btn-number` — 숫자 (보라 계열)
- `.btn-operator` — 연산자 (핑크/레드 계열)
- `.btn-special` — AC, +/-, % (파랑 계열)
- `.btn-equals` — = (강조 핑크)

## 스타일 규칙

- 폰트: Pretendard — CDN `https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css`
- `font-family: 'Pretendard', -apple-system, sans-serif;`

## 응답 규칙

- 모든 결과값과 설명은 반드시 한글로 작성한다.
- 작업 시작 전 항상 Plan Mode로 계획을 먼저 수립한다.
