# Cosmetics.lk Product Scraper (GitHub Actions)

Opens the **Women** collection on cosmetics.lk, visits each product one by one, and saves
**image, product name, price and description** (plus brand, type, sale price, SKU and links)
to an Excel file. The product image is embedded directly in each row.

## Files

```
.github/workflows/scrape.yml   <- the GitHub Actions workflow
scraper.py                     <- the scraper
requirements.txt               <- Python packages
```

## Setup (one time)

1. Create a new GitHub repository (public or private).
2. Upload these files keeping the same folder structure
   (the `.github/workflows/` folder matters).
3. Go to **Settings → Actions → General → Workflow permissions** and select
   **Read and write permissions** (only needed if you want the "commit to repo" option).

## Run it

1. Open the **Actions** tab → **Scrape cosmetics.lk products** → **Run workflow**.
2. Leave `max_products` at **10** for the demo (or change it), click **Run workflow**.
3. When the run finishes (≈1 minute for 10 products), open the run and download
   **cosmetics-products-N** under *Artifacts*. Inside is `cosmetics_products.xlsx`.

It also runs automatically every day at 06:00 Sri Lanka time. Delete the `schedule:` lines
in `scrape.yml` if you don't want that.

## Run locally

```bash
pip install -r requirements.txt
python scraper.py --max-products 10
# output/cosmetics_products.xlsx
```

Options: `--collection men`, `--max-products 100`, `--output myfile.xlsx`.

## Excel columns

`# | Image | Product Name | Brand | Type | Price (Rs) | Original Price (Rs) | On Sale | SKU | Description | Product URL | Image URL`

## Notes

- The site runs on Shopify, so the scraper reads each product's JSON version
  (`/products/<name>.json`), which is more reliable than reading the page layout.
  If that ever fails it falls back to reading the HTML page.
- There is a 1-second pause between requests so the site isn't overloaded.
- Check the site's terms of use before scraping large volumes or reusing the content commercially.
