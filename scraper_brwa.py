import urllib.request
import re
import json
import time
import sys
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://brwa.or.id/"

def fetch_page(limit=100, page=1):
    url = f"https://brwa.or.id/wa/index/0/{limit}/{page}?q=0&p=0&k=0"
    req = urllib.request.Request(
        url,
        headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"Error fetching page {page}: {e}")
        return None

def parse_html_table(html):
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table', {'class': 'table-hover'})
    if not table:
        return [], 0, 0

    # Extract total pages from pagination links if present
    total_pages = 1
    pagination = soup.find('ul', {'class': 'pagination'})
    if pagination:
        page_links = pagination.find_all('a')
        for link in page_links:
            href = link.get('href', '')
            match = re.search(r'/(\d+)\?q=', href)
            if match:
                pg_num = int(match.group(1))
                if pg_num > total_pages:
                    total_pages = pg_num

    tbody = table.find('tbody')
    if not tbody:
        return [], total_pages, 0

    rows = tbody.find_all('tr')
    extracted_data = []

    for tr in rows:
        tds = tr.find_all('td')
        if len(tds) < 10:
            continue

        no = tds[0].get_text(strip=True)
        
        # Detail URL
        detail_a = tds[1].find('a') or tds[3].find('a')
        detail_href = detail_a.get('href', '') if detail_a else ''
        if detail_href and not detail_href.startswith('http'):
            url_detail = BASE_URL + detail_href.lstrip('/')
        else:
            url_detail = detail_href
            
        wa_id = detail_href.split('/')[-1] if detail_href else ''

        tanggal_daftar = tds[2].get_text(strip=True)
        nama_wilayah_adat = tds[3].get_text(strip=True)
        provinsi = tds[4].get_text(strip=True)
        kab_kota = tds[5].get_text(strip=True)
        kecamatan = tds[6].get_text(strip=True)

        # Peta checkmark
        peta_icon = tds[7].find('i', {'class': re.compile(r'fa-check')})
        peta = "Ada" if peta_icon else tds[7].get_text(strip=True)
        if not peta or peta == "-":
            peta = "Tidak Ada" if not peta_icon else "Ada"

        # Status badge
        status_span = tds[8].find('span')
        status = status_span.get_text(strip=True) if status_span else tds[8].get_text(strip=True)

        # Kebijakan checkmark
        kebijakan_icon = tds[9].find('i', {'class': re.compile(r'fa-check')})
        kebijakan = "Ada" if kebijakan_icon else tds[9].get_text(strip=True)
        if kebijakan == "-":
            kebijakan = "Tidak Ada"

        extracted_data.append({
            'no': int(no) if no.isdigit() else no,
            'id_wa': wa_id,
            'tanggal_daftar': tanggal_daftar,
            'nama_wilayah_adat': nama_wilayah_adat,
            'provinsi': provinsi,
            'kabupaten_kota': kab_kota,
            'kecamatan': kecamatan,
            'peta': peta,
            'status': status,
            'kebijakan': kebijakan,
            'url_detail': url_detail
        })

    return extracted_data, total_pages, len(rows)

def main():
    print("=== SCRAPER INDEKS WILAYAH ADAT BRWA ===")
    print("Sedang memeriksa jumlah data...")

    limit = 100
    first_page_html = fetch_page(limit=limit, page=1)
    if not first_page_html:
        print("Gagal mengambil halaman pertama.")
        sys.exit(1)

    initial_data, total_pages, rows_in_first = parse_html_table(first_page_html)
    print(f"Halaman 1 (Limit {limit}): Ditemukan {rows_in_first} data. Total Halaman Est: {total_pages}")

    all_records = list(initial_data)

    for page in range(2, total_pages + 1):
        print(f"Mengambil Halaman {page}/{total_pages}...", end="", flush=True)
        html = fetch_page(limit=limit, page=page)
        if html:
            data, _, count = parse_html_table(html)
            all_records.extend(data)
            print(f" Sukses ({count} data)")
        else:
            print(" Gagal")
        time.sleep(0.3)

    print(f"\nTotal Keseluruhan Data Terkumpul: {len(all_records)} record")

    if not all_records:
        print("Tidak ada data yang berhasil diambil.")
        return

    # Convert to DataFrame
    df = pd.DataFrame(all_records)

    # Save to CSV
    csv_file = "data_brwa_indeks.csv"
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"Saved: {csv_file}")

    # Save to Excel
    excel_file = "data_brwa_indeks.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')
    print(f"Saved: {excel_file}")

    # Save to JSON
    json_file = "data_brwa_indeks.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)
    print(f"Saved: {json_file}")

    print("\nSample 5 Data Teratas:")
    print(df.head())

if __name__ == "__main__":
    main()
