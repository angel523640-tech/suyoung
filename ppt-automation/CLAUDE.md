# PPT Automation

회사 공식 템플릿(`template/company_template.pptx`)과 입력 컨텐츠(`input/content.md`)를 기반으로 PPT 보고서를 자동 생성하여 `output/`에 저장합니다.

## 폴더 구조

```
/ppt-automation
├── CLAUDE.md                       # 프로젝트 프롬프트 / 지침
├── /template
│   └── company_template.pptx       # 회사 공식 템플릿
├── /input
│   └── content.md                  # 매번 바뀌는 컨텐츠 입력
└── /output
    └── (생성된 보고서가 저장됨)
```

## 사용 방법

1. `template/company_template.pptx`에 회사 공식 템플릿 파일을 배치합니다.
2. `input/content.md`에 이번 보고서에 넣을 컨텐츠를 작성합니다.
3. Claude에게 실행을 요청하면 `output/`에 보고서가 생성됩니다.

## 프롬프트

> 이 섹션에 "아래 프롬프트"로 사용할 내용을 붙여넣어 주세요.
