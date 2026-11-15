# Web_scraper_faster
beder version of my last work its much faster and when downloading images its only a little slower the first time

Hi! I'm an IT high school senior and this is my second Python automation project. This script is a highly optimized, asynchronous web scraper designed to extract product data, handle complex web layouts, and generate beautiful, production-ready Excel catalogs with embedded product images.

For this project, I chose a popular alternative clothing e-shop (**Metalshop**) to test my skills on a completely different website structure, using advanced optimizations to make it as fast and efficient as possible.

## 🚀 Key Features & Performance Optimizations
* **Blazing Fast Headless Scraping:** Runs completely in the background (`headless=True`) and uses aggressive request routing (`page.route`) to block heavy CSS stylesheets and fonts. This reduces page loading overhead by up to 90%.
* **Smart Content Loading:** Implements `wait_until="domcontentloaded"` so the script doesn't waste time waiting for tracking scripts and analytics to load before collecting data.
* **Local Image Caching System:** Before initiating any network download, the script uses `os.path.exists()` to check if the product image is already saved locally. If it is, the download is skipped, making subsequent runs **over 50% faster** and saving bandwidth!
* **Bulletproof URL Encoding:** Features an `encode_url_for_request` component that dynamically normalizes and signs messy image URLs (handling query parameters, safe character mapping, and missing protocols) before sending requests.
* **Professional Typography & Conditional Fills:** Built with `openpyxl` using crisp **Calibri 14pt** typography. The header features a dark anthracite theme with contrasting white text, and product names are styled in bold. Entire rows are dynamically color-coded with subtle pastel fills based on their availability and marketing state (`Bestseller -> Discount -> New -> In stock -> Damaged -> Unavailable`).
* **Clean Layout & Border Isolation:** Column A auto-wraps long product titles to keep the layout scannable. Miniatures (140x100px) are cleanly embedded into Column E, which has borders explicitly disabled to maintain a flawless, modern corporate catalog aesthetic.

## 💻 Tech Stack
* **Python 3** (Core logic)
* **Playwright** (Fast browser automation & async network interception)
* **Pandas** (Data wrangling, duplicate removal & sorting priority pipelines)
* **OpenPyXL** (Advanced spreadsheet manipulation, formatting & rendering graphics)

## 📊 Output Preview
The pipeline generates a beautifully formatted spreadsheet (`results_Tshirts.xlsx`) with the following structure:
`Product Name` (Bold, auto-wrap) | `State` (Underlined) | `Price with VAT` | `Price after discount` | `Front image` (Clean border-free embedded cell graphics)
