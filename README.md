# StockSense 抽圖後端

接收 Excel,抽出內嵌圖片並對應款號,上傳到 Supabase Storage,回傳結果供前端核對。

## 環境變數(在 Render 設定)

|變數                    |值                                                         |
|----------------------|----------------------------------------------------------|
|`SUPABASE_URL`        |`https://fuuinvuimfrhaddmkjbw.supabase.co`                |
|`SUPABASE_SERVICE_KEY`|(從 Supabase Settings → API → service_role secret 複製)      |
|`SUPABASE_JWT_SECRET` |(從 Supabase Settings → API → JWT Settings → JWT Secret 複製)|
|`ALLOWED_ORIGINS`     |`https://stocksense-app-eta.vercel.app` (你的前端網址)          |

## 部署到 Render

1. 把這個資料夾推到一個 GitHub repo(例如 `stocksense-backend`)
1. Render → New → Web Service → 選你的 repo
1. Runtime: Docker
1. Plan: Starter ($7/月)
1. 加上上面四個環境變數
1. Deploy

## 端點

- `GET /` → 健康檢查
- `POST /extract` → 抽圖
  - Header: `Authorization: Bearer <access_token>`
  - Body: multipart/form-data,欄位 `file` = .xls 或 .xlsx 檔
  - 回傳: `{ ok, extracted, uploaded, results: [{ style_no, image_url, filename }] }`