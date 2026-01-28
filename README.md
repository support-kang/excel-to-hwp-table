# ExcelToHwp (엑셀 to 한글 변환기)

엑셀 파일(.xlsx, .xls)의 데이터를 읽어 **한글(.hwp)** 파일로 변환해주는 Java 애플리케이션입니다.
단순 텍스트 나열이 아닌, **한글 표(Table)** 형식으로 깔끔하게 변환됩니다.

## 📋 주요 기능

- **GUI 지원**: 직관적인 사용자 인터페이스 제공.
- **엑셀 파일 선택**: `.xlsx` 및 `.xls` 파일 지원.
- **시트 및 범위 지정**: 특정 시트의 원하는 행/열 범위만 선택하여 변환 가능.
- **자동 표 생성**: 엑셀 데이터를 기반으로 한글 표(ControlTable) 자동 생성.
- **EXE 실행 파일 생성**: 윈도우 사용자를 위한 `.exe` 실행 파일 빌드 지원.

## 🛠️ 요구 사항 (Prerequisites)

- **Java Development Kit (JDK) 8 이상**
- **Maven 3.x**

## 🚀 빌드 및 실행 방법

### 1. 프로젝트 빌드

프로젝트 루트(`ExcelToHwpJava`)에서 터미널을 열고 아래 명령어를 실행하세요.

```bash
mvn package
```

### 2. 실행

빌드가 성공하면 `target` 폴더에 실행 파일들이 생성됩니다.

#### JAR로 실행 (모든 OS)
```bash
java -jar target/ExcelToHwp-1.0-SNAPSHOT-shaded.jar
```

#### EXE로 실행 (Windows)
`target` 폴더 내의 `ExcelToHwp.exe` 파일을 더블 클릭하여 실행합니다.

## 📱 사용 방법

1. **Excel File** 버튼을 눌러 변환할 엑셀 파일을 선택합니다.
2. (선택 사항) **Sheet Name**에 변환할 시트 이름을 입력합니다. (비워두면 첫 번째 시트 사용)
3. (선택 사항) **Start/End Row/Col**에 변환할 데이터 범위를 지정합니다.
4. **Convert to HWP** 버튼을 클릭합니다.
5. 변환이 완료되면 원본 엑셀 파일과 같은 위치에 `_Result.hwp` 파일이 생성됩니다.

## 📁 프로젝트 구조

```
ExcelToHwpJava/
├── src/
│   └── main/java/com/example/ExcelToHwp.java  # 메인 소스 코드
├── pom.xml                                    # Maven 설정 (의존성 및 빌드 플러그인)
├── .gitignore                                 # Git 무시 설정
└── README.md                                  # 프로젝트 설명
```

## 📝 라이브러리 정보

이 프로젝트는 다음 오픈 소스 라이브러리를 사용합니다.
- **hwplib (v1.1.9)**: 한글(HWP) 파일 생성 및 조작.
- **Apache POI (v5.2.5)**: 엑셀 파일 읽기.
- **Launch4j**: EXE 실행 파일 래핑.
