# Cellendar

PC의 `캘린더(자동).xlsx`와 iPhone/Android PWA 캘린더를 사용자별 OneDrive를 통해 동기화하는 GitHub Pages 앱입니다.

## 구조

```text
공용 GitHub Pages 앱
├─ 사용자 A 로그인 → A의 OneDrive/Cellendar/캘린더(자동).xlsx
├─ 사용자 B 로그인 → B의 OneDrive/Cellendar/캘린더(자동).xlsx
└─ 사용자 C 로그인 → C의 OneDrive/Cellendar/캘린더(자동).xlsx
```

GitHub에는 앱 코드만 배포됩니다. 일정 및 Excel 파일은 사용자의 Microsoft 저장소에 남습니다.

## 산출물

- `app/`: GitHub Pages에 배포되는 PWA
- `outputs/cellendar/캘린더(자동).xlsx`: PC/OneDrive용 동기화 엑셀
- `docs/관리자_배포안내.md`: Entra 및 GitHub Pages 설정
- `docs/사용자_설치안내.md`: iPhone/Android 설치 방법

## 로컬 확인

```powershell
python -m http.server 8080
```

`http://localhost:8080/app/`에서 확인합니다. Microsoft 로그인을 시험하려면 이 URI도 Entra의 SPA 리디렉션 URI에 등록해야 합니다.

