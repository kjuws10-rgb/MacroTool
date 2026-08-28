# MacroTool

보안정책을 준수하는 업무용 매크로와 배포 문서를 관리하는 저장소입니다.

## 포함된 도구

### Secure Excel Transfer VBA

관리자가 승인한 폴더의 `.xlsx` 또는 `.csv` 데이터를 정확히 승인된 `.xlsm` 문서의 빈 범위로 값만 가져옵니다.

- 클립보드, 강제 붙여넣기, 키 입력 자동화 사용 안 함
- DRM/DLP 차단 시 우회 없이 작업 중단
- 승인 상태·만료일·대상 전체 경로·처리 한도 검증
- 원본 파일명, 대상 파일명, 처리 시각, 행 수만 감사 기록
- 설치 및 보안 검토 절차 포함

자세한 내용은 [`SecureExcelTransferVBA/README.md`](SecureExcelTransferVBA/README.md)를 참고하십시오.

## 개발 방식

변경 사항은 기능 브랜치에서 구현·검증한 뒤 Pull Request를 통해 `main`으로 병합합니다. 세부 규칙은 [`CONTRIBUTING.md`](CONTRIBUTING.md)에 정리되어 있습니다.
