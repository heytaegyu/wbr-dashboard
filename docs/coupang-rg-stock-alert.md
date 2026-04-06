# Coupang Rocket Growth Daily Slack Report

이 구성은 쿠팡 로켓그로스 3개 SKU를 매일 점검해서 Slack으로 보내는 용도입니다.

현재 기본 감시 SKU는 [config/coupang-rg-alert.json](/Users/taegyu/Documents/New%20project/config/coupang-rg-alert.json) 에 들어 있습니다.

- `65272295` : 1개입
- `66370744` : 3개입
- `66370743` : 5개입

현재 품절임박 기준은 `판매가능 예상일수 28일 이하` 입니다.

## Slack에 가는 내용

1. 품절임박 여부
2. 3개 SKU 전체 재고 현황
3. 판매 분석
   - 어제 판매수량 / 매출
   - 오늘 `09:00 KST` 기준 판매수량 / 매출
4. 매출 / 비용 / 이익
   - 원가표 기준 추정값
   - 제품원가, 입출고+배송비, 광고비, 판매수수료, 선정산, 부가세를 합산해서 추정 이익 계산

## 데이터 출처

- 재고: 쿠팡 `로켓창고 재고 API`
- 상품명 매핑: 쿠팡 `상품 목록 페이징 조회 API`
- 판매량/매출: 쿠팡 `RG Order API`
- 비용/이익: `config/coupang-rg-alert.json` 에 넣은 원가표 기준 추정

중요: 비용/이익은 쿠팡 정산탭 실시간 값이 아니라, 원가표를 기준으로 계산한 추정값입니다.

## 필요한 값

- `COUPANG_ACCESS_KEY`
- `COUPANG_SECRET_KEY`
- `COUPANG_VENDOR_ID`
- `SLACK_WEBHOOK_URL`

## 로컬 실행

```bash
export COUPANG_ACCESS_KEY="..."
export COUPANG_SECRET_KEY="..."
export COUPANG_VENDOR_ID="A00012345"
export SLACK_WEBHOOK_URL="https://hooks.slack.com/services/..."

python3 scripts/check_coupang_rg_inventory.py --send-healthy-summary
```

실제 전송 없이 메시지 내용만 보려면:

```bash
python3 scripts/check_coupang_rg_inventory.py --send-healthy-summary --dry-run
```

## 설정 파일

[config/coupang-rg-alert.json](/Users/taegyu/Documents/New%20project/config/coupang-rg-alert.json) 에서 아래를 조정할 수 있습니다.

- `days_threshold`
- `snapshot_hour_kst`
- SKU 목록
- SKU별 원가표

## GitHub Actions

PC가 꺼져 있어도 돌리려면 `.github/workflows/coupang-rg-stock-alert.yml` 을 사용하면 됩니다.

- 스케줄: 매일 `00:05 UTC`
- 한국 시간 기준: 매일 `09:05 KST`

GitHub Secrets

- `COUPANG_ACCESS_KEY`
- `COUPANG_SECRET_KEY`
- `COUPANG_VENDOR_ID`
- `SLACK_WEBHOOK_URL`

GitHub Variables

- `COUPANG_LOW_STOCK_QTY` 선택
- `COUPANG_LOW_STOCK_DAYS` 선택
- `COUPANG_ALERT_MAX_ITEMS` 선택
- `COUPANG_TARGET_SKUS` 선택
- `COUPANG_SNAPSHOT_HOUR_KST` 선택

## 참고한 공식 문서

- 로켓창고 재고 API:
  https://developers.coupangcorp.com/hc/ko/articles/41090779386521-%EB%A1%9C%EC%BC%93%EC%B0%BD%EA%B3%A0-%EC%9E%AC%EA%B3%A0-API
- 상품 목록 페이징 조회:
  https://developers.coupangcorp.com/hc/en-us/articles/39427498030745-Product-list-paging-query-Rocket-Growth-Rocket-Growth-Marketplace-Hybrid-Products
- RG Order API:
  https://developers.coupangcorp.com/hc/en-us/articles/41131195825433-RG-Order-API-List-Query
- 매출내역 조회:
  https://developers.coupangcorp.com/hc/ko/articles/360033922413-%EB%A7%A4%EC%B6%9C%EB%82%B4%EC%97%AD-%EC%A1%B0%ED%9A%8C
- HMAC 서명:
  https://developers.coupangcorp.com/hc/en-us/articles/360033461914-Creating-HMAC-Signature-
