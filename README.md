# 해외주식 양도세 프로그램

## 소개

* 홈페이지
  - [https://sdrlurker.notion.site/166d1dff257980d3a8e8c71e5c022e81](https://sdrlurker.notion.site/166d1dff257980d3a8e8c71e5c022e81)
* 증권사의 해외주식 양도세 자료를 웹브라우저 크롬 selenium을 통해 스크래핑하여 홈택스로 제출할 수 있는 엑셀 파일을 생성할 수 있습니다.
* 가능한 증권사 (2025.04.02)
  - 한국투자증권
  - 삼성증권
  - 키움증권
  - 신한투자증권
  - 하나증권

## 설치방법

* 의존성 모듈 설치

```shell
pip install -r requirements.txt
```

* 윈도우즈 실행파일 생성

```shell
show.bat (upx 디렉터리)
```

* 맥 실행파일 생성

```shell
./show.sh
```

## 소스설명

* show.py
  - 엔트리 포인트. 프로그램의 진입점.
  - kivy 모듈을 통해 사용자가 사용하는 화면(UI)을 만듭니다.
* collect.py
  - 크롬 selenium을 통해 증권사 홈페이지에서 해외주식 양도세 자료를 스크래핑 합니다.
  - 증권사 별로 class로 구성됨.
