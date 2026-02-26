#!/usr/bin/env python3
"""
Weekly Consulting Job Board Scraper
Scrapes jobs from MyConsultingOffer and Management Consulted job boards,
then generates a McKinsey-styled Excel file.

Requirements: playwright, openpyxl
Install: pip install playwright openpyxl && playwright install chromium
"""

import re, json, time, os
from datetime import date, datetime
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


MCO_URL = "https://jobs.myconsultingoffer.org/jobs"
MC_URL = "https://jobs.managementconsulted.com/jobs"
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", ".")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"consulting_job_board_{date.today().isoformat()}.xlsx")

COMPANY_WEBSITES = {
    'McKinsey & Company': 'https://www.mckinsey.com',
    'Boston Consulting Group (BCG)': 'https://www.bcg.com',
    'Boston Consulting Group': 'https://www.bcg.com',
    'Bain & Company': 'https://www.bain.com',
    'Roland Berger': 'https://www.rolandberger.com',
    'Oliver Wyman': 'https://www.oliverwyman.com',
    'Alvarez & Marsal': 'https://www.alvarezandmarsal.com',
    'EY': 'https://www.ey.com',
    'EY-Parthenon': 'https://www.ey.com/en_gl/services/strategy/ey-parthenon',
    'Deloitte': 'https://www.deloitte.com',
    'KPMG UK': 'https://www.kpmg.co.uk',
    'KPMG Canada': 'https://home.kpmg/ca',
    'KPMG US': 'https://www.kpmg.us',
    'PwC': 'https://www.pwc.com',
    'PwC Canada': 'https://www.pwc.com/ca',
    'Accenture': 'https://www.accenture.com',
    'Capital One': 'https://www.capitalone.com',
    'Google': 'https://careers.google.com',
    'Guidehouse': 'https://guidehouse.com',
    'Huron': 'https://www.huronconsultinggroup.com',
    'Simon-Kucher': 'https://www.simon-kucher.com',
    'FTI Consulting': 'https://www.fticonsulting.com',
    'Booz Allen Hamilton': 'https://www.boozallen.com',
    'L.E.K. Consulting': 'https://www.lek.com',
    'Kearney': 'https://www.kearney.com',
    'AlixPartners': 'https://www.alixpartners.com',
    'Infosys': 'https://www.infosys.com',
    'Cognizant': 'https://www.cognizant.com',
    'IQVIA': 'https://www.iqvia.com',
    'Visa': 'https://www.visa.com',
    'The Walt Disney Company': 'https://thewaltdisneycompany.com',
    'American Express': 'https://www.americanexpress.com',
    'Ford Motor Company': 'https://www.ford.com',
    'General Motors': 'https://www.gm.com',
    'Pfizer': 'https://www.pfizer.com',
    'The Carlyle Group': 'https://www.carlyle.com',
    'RBC': 'https://www.rbc.com',
    'Trinity Life Sciences': 'https://trinitylifesciences.com',
    'Mercer': 'https://www.mercer.com',
    'West Monroe': 'https://www.westmonroe.com',
    'Stax': 'https://www.stax.com',
    'Porsche Consulting': 'https://www.porsche-consulting.com',
    'RSM Canada': 'https://www.rsmcanada.com',
    'The Coca-Cola Company': 'https://www.coca-colacompany.com',
    'Kantar': 'https://www.kantar.com',
    'ClearView Healthcare Partners': 'https://www.clearviewhcp.com',
    'Clarkston Consulting': 'https://clarkstonconsulting.com',
    'North Highland': 'https://www.northhighland.com',
}


def scrape_jboard_site(page, site_url, site_name):
    print(f"\nScraping {site_name}: {site_url}")
    page.goto(site_url, wait_until="networkidle", timeout=60000)
    time.sleep(3)

    prev_count = 0
    stale_rounds = 0
    while stale_rounds < 5:
        count = page.evaluate("document.querySelectorAll('a[href*=/jobs/]').length")
        if count == prev_count:
            stale_rounds += 1
        else:
            stale_rounds = 0
            prev_count = count
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1.5)

    print(f"Found ~{prev_count} job elements on page")

    jobs = page.evaluate("""() => {
        const jobs = [];
        const seen = new Set();
        const jobLinks = document.querySelectorAll('a[href*="/jobs/"]');
        for (const link of jobLinks) {
            const href = link.getAttribute('href') || '';
            const match = href.match(/\\/jobs\\/(\\d+)/);
            if (!match) continue;
            const jobId = match[1];
            if (seen.has(jobId) || jobId.length < 6) continue;
            seen.add(jobId);
            const card = link.closest('[class*=job], [class*=Job], li, article, div') || link;
            const text = card.innerText || '';
            const lines = text.split('\\n').map(l => l.trim()).filter(Boolean);
            const title = lines[0] || '';
            const company = lines[1] || '';
            let location = '', jobType = '', category = '';
            for (const line of lines.slice(2)) {
                if (/full.time|part.time|contract|internship/i.test(line) && !jobType) jobType = line;
                else if (/,/.test(line) && /(US|CA|GB|DE|FR|AU|IN|SG|AE|NL|CH|BR|MX|JP|KR|ES|IT|PT|BE|SE|PL|RO|HU|AT)$/i.test(line.split(',').pop().trim())) location = line;
                else if (/consulting|strategy|advisory|corporate/i.test(line) && !category) category = line;
            }
            if (title && title.length > 3) jobs.push({title, company, jobType, category, location, jobId});
        }
        return jobs;
    }""")

    print(f"Extracted {len(jobs)} jobs from {site_name}")
    return jobs


def scrape_jboard_api(page, base_url):
    print(f"Attempting API-based extraction from {base_url}...")
    all_jobs = []
    api_responses = []

    def handle_response(response):
        if 'jboard-api' in response.url or '/api/' in response.url:
            try:
                data = response.json()
                if isinstance(data, list): api_responses.extend(data)
                elif isinstance(data, dict) and 'data' in data: api_responses.extend(data['data'])
            except: pass

    page.on("response", handle_response)
    page.goto(base_url, wait_until="networkidle", timeout=60000)
    for _ in range(50):
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1)
    page.remove_listener("response", handle_response)

    if api_responses:
        print(f"Captured {len(api_responses)} jobs via API")
        for item in api_responses:
            if isinstance(item, dict) and 'title' in item:
                all_jobs.append({
                    'title': item.get('title', ''),
                    'company': item.get('company_name', item.get('company', '')),
                    'jobType': item.get('employment_type', item.get('type', '')),
                    'category': item.get('category', ''),
                    'location': item.get('location', ''),
                    'jobId': str(item.get('id', item.get('slug', ''))),
                })
    return all_jobs


def scrape_site(page, site_url, site_name):
    jobs = scrape_jboard_site(page, site_url, site_name)
    if len(jobs) > 50: return jobs
    print(f"DOM scraping got only {len(jobs)} jobs, trying API approach...")
    api_jobs = scrape_jboard_api(page, site_url)
    return api_jobs if len(api_jobs) > len(jobs) else jobs


def build_excel(all_jobs, output_path):
    print(f"\nBuilding Excel with {len(all_jobs)} jobs...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Job Board"

    MCK_NAVY, MCK_BLUE = '003A5C', '0072CE'
    MCK_LIGHT_BLUE, MCK_LIGHT_GRAY = 'E8F4FD', 'F5F5F5'
    MCK_WHITE, MCK_BORDER_COLOR = 'FFFFFF', 'D0D0D0'

    header_fill = PatternFill('solid', fgColor=MCK_NAVY)
    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data_font = Font(name='Arial', size=9.5, color='333333')
    link_font = Font(name='Arial', size=9.5, color=MCK_BLUE, underline='single')
    data_align = Alignment(vertical='center', wrap_text=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color=MCK_BORDER_COLOR),
        right=Side(style='thin', color=MCK_BORDER_COLOR),
        top=Side(style='thin', color=MCK_BORDER_COLOR),
        bottom=Side(style='thin', color=MCK_BORDER_COLOR),
    )
    alt_fill = PatternFill('solid', fgColor=MCK_LIGHT_GRAY)
    white_fill = PatternFill('solid', fgColor=MCK_WHITE)

    ws.merge_cells('A1:H1')
    ws['A1'].value = 'MBB Prep - Management Consulting Job Board'
    ws['A1'].font = Font(name='Arial', bold=True, color='FFFFFF', size=14)
    ws['A1'].fill = PatternFill('solid', fgColor=MCK_BLUE)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    ws.merge_cells('A2:H2')
    ws['A2'].value = f'{len(all_jobs)} consulting jobs  |  Updated {date.today().strftime("%b %d, %Y")}'
    ws['A2'].font = Font(name='Arial', size=10, color='666666', italic=True)
    ws['A2'].fill = PatternFill('solid', fgColor=MCK_LIGHT_BLUE)
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 8

    headers = ['#', 'Job Title', 'Firm', 'Category', 'Job Type', 'Location', 'Firm Website', 'Application Link']
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws.row_dimensions[4].height = 30

    for i, job in enumerate(all_jobs):
        row = i + 5
        row_fill = alt_fill if i % 2 == 0 else white_fill
        ws.cell(row=row, column=1, value=i+1).font = Font(name='Arial', size=9, color='999999')
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=2, value=job['title']).font = Font(name='Arial', size=9.5, color='333333', bold=True)
        ws.cell(row=row, column=2).alignment = data_align
        ws.cell(row=row, column=3, value=job['company']).font = data_font
        ws.cell(row=row, column=3).alignment = data_align
        cat = job.get('category', '') or 'Management Consulting'
        ws.cell(row=row, column=4, value=cat).font = data_font
        ws.cell(row=row, column=4).alignment = center_align
        ws.cell(row=row, column=5, value=job.get('jobType', '')).font = data_font
        ws.cell(row=row, column=5).alignment = center_align
        ws.cell(row=row, column=6, value=job.get('location', '')).font = data_font
        ws.cell(row=row, column=6).alignment = data_align
        website = job.get('website', '')
        if website:
            domain = website.replace('https://', '').replace('http://', '').split('/')[0]
            c = ws.cell(row=row, column=7, value=domain)
            c.font = link_font
            c.hyperlink = website
        ws.cell(row=row, column=7).alignment = data_align
        apply_url = job.get('applyLink', '')
        if apply_url:
            c = ws.cell(row=row, column=8, value='Apply')
            c.font = link_font
            c.hyperlink = apply_url
        ws.cell(row=row, column=8).alignment = data_align
        for col in range(1, 9):
            ws.cell(row=row, column=col).fill = row_fill
            ws.cell(row=row, column=col).border = thin_border

    widths = {'A': 5, 'B': 55, 'C': 28, 'D': 22, 'E': 12, 'F': 35, 'G': 30, 'H': 12}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w
    ws.auto_filter.ref = f'A4:H{4 + len(all_jobs)}'
    ws.freeze_panes = 'A5'
    wb.save(output_path)
    print(f"Saved: {output_path}")


def main():
    all_jobs = []
    seen_ids = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1280, "height": 900},
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
        )

        page_mco = context.new_page()
        mco_jobs = scrape_site(page_mco, MCO_URL, "MyConsultingOffer")
        for j in mco_jobs:
            jid = j.get('jobId', '')
            if jid and jid not in seen_ids:
                seen_ids.add(jid)
                j['website'] = COMPANY_WEBSITES.get(j['company'], '')
                j['applyLink'] = f"https://jobs.myconsultingoffer.org/jobs/{jid}/apply"
                all_jobs.append(j)
        page_mco.close()

        page_mc = context.new_page()
        mc_jobs = scrape_site(page_mc, MC_URL, "Management Consulted")
        for j in mc_jobs:
            jid = j.get('jobId', '')
            if jid and jid not in seen_ids:
                seen_ids.add(jid)
                j['website'] = COMPANY_WEBSITES.get(j['company'], '')
                j['applyLink'] = f"https://jobs.managementconsulted.com/jobs/{jid}/apply"
                all_jobs.append(j)
        page_mc.close()
        browser.close()

    print(f"\nTotal unique jobs: {len(all_jobs)}")
    if all_jobs:
        build_excel(all_jobs, OUTPUT_FILE)
    else:
        print("ERROR: No jobs scraped. Check if site structure changed.")
        exit(1)


if __name__ == "__main__":
    main()
