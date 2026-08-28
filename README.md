# MacroTool

보안정책을 준수하는 업무용 매크로와 배포 문서를 관리하는 저장소입니다.

## 포함된 도구

### Secure Excel Transfer VBA

관리자가 승인한 폴더의 `.xlsx` 또는 `.csv` 데이터를 정확히 승인된 `.xlsm` 문서의 빈 범위로 값만 가져옵니다.

- 클립보드, 강제 붙여넣기, 키 입력 자동화 사용 안 함
- 전용 시작 화면과 큰 버튼, 마우스 범위 선택 제공
- DRM/DLP 차단 시 우회 없이 작업 중단
- 승인 상태·만료일·대상 전체 경로·처리 한도 검증
- 원본 파일명, 대상 파일명, 처리 시각, 행 수만 감사 기록
- 설치 및 보안 검토 절차 포함

자세한 내용은 [`SecureExcelTransferVBA/README.md`](SecureExcelTransferVBA/README.md)를 참고하십시오.

## 저장소 업데이트

Git clone으로 받은 저장소의 루트에서 다음 파일을 실행하면 현재 브랜치의 upstream을 안전하게 가져옵니다.

```bat
pull.bat
```

- `pull.bat -Prune`: 원격에서 삭제된 추적 브랜치도 함께 정리
- `pull.bat -Help`: 사용법 표시
- 커밋되지 않은 추적 파일이 있으면 중단
- 자동 `stash`, `reset`, 파일 삭제 없이 `git pull --ff-only`만 수행
- PowerShell은 해당 프로세스에만 `RemoteSigned`를 요청하며 조직 그룹 정책은 그대로 우선 적용

ZIP으로 받은 폴더에서는 사용할 수 없습니다. 최초 한 번은 `git clone https://github.com/kjuws10-rgb/MacroTool.git`으로 받아야 합니다.

## 개발 방식

변경 사항은 기능 브랜치에서 구현·검증한 뒤 Pull Request를 통해 `main`으로 병합합니다. 세부 규칙은 [`CONTRIBUTING.md`](CONTRIBUTING.md)에 정리되어 있습니다.
