[https://sidharth005p.github.io/chennai-ecom-status](https://sidharth005p.github.io/chennai-ecom-status)

# Chennai & ROTN E-Com Review Status (GitHub Pages)

Files
- index.html – public page with the original Review dashboard, a separate Dashboard based on new rules, Rejected tickets, Rejection insights and If all tickets accepted
- admin.html – for the branch: reads the daily Excel files and makes the new data files
- core.js – shared calculation code used by the admin and public pages; upload it when updated
- data.json – Chennai / ROTN review tickets (ticket ID, branch, date; replace daily)
- report.json – calculations and summaries shown by the dashboards (replace daily)

Set up once *UNCONVENTIONAL DO NOT FOLLOW*
1. Create a free GitHub account, then New repository: name chennai-ecom-status, Public, tick "Add a README".
2. Add file > Upload files: drag index.html, admin.html, core.js, data.json and report.json, then Commit changes.
3. Settings > Pages > Source: Deploy from a branch, Branch main, folder / (root), Save.
4. After 1-2 minutes the site is live at https://YOUR-USERNAME.github.io/chennai-ecom-status/ and the admin page at .../admin.html

Daily update
1. Open .../admin.html.
2. Choose the daily review Excel (Branch, Ticket ID, Added Time) and click Download data.json. The file can cover several months (for example 19 May till today). Every Chennai / ROTN ticket in it, from any month, counts as a review (no Order ID checking).
3. In the second box choose the reduced calls file (all INSTALLATION calls of the last 6 months, every dealer and status, only the needed columns; keep it under 45 MB). The rejected-ticket reasons need the cancelled / other-dealer calls to be in this file. If the file has a call date column, a "Month for the review %" list appears (current month is selected automatically); then click Download report.json.
   (If you skip step 2 the report uses the review data already on the website.)
4. In the repository click Add file > Upload files, drop data.json and report.json (names must stay exactly the same), Commit changes.
5. Wait about a minute, then refresh the main page (Ctrl+F5).

How the report is calculated
- E-commerce installation ticket = Call Type "INSTALLATION CALL" with Dealer Name containing "Flipkart India" or "Cloudtail India".
- Closed = status Closed or Resolved (soft closure). The admin page has a tick box to count only Closed.
- Original Review dashboard: a review is any ticket present in the uploaded review data (data.json, any of the 6 months), whatever its Order ID.
- Dashboard based on new rules: only review tickets whose Added Time is in the selected report month or the month immediately before it are eligible. Other and undated uploads are excluded from both accepted and rejected counts. For a November report, this means November and October uploads.
- Denominator = e-commerce installation tickets closed in the chosen (current) month only. Numerator = every ticket of the review file that matches a closed e-commerce installation ticket of ANY month in the calls file (the review file timestamp is only the upload time, the ticket itself can belong to any of the 6 months). A technician or franchise with more reviews than current-month closed tickets can therefore show above 100%.
- Only these products are counted (Product column): WM, WD = Washing Machine; CD = Cloth Dryer; MW = Microwave -> Category 1. AC = Air Conditioner; RF = Refrigerator; DW = Dishwasher -> Category 2. IND, KF and any other product are ignored.
- Slab is worked out separately for each category: reviews / closed tickets of that technician in that category = 25, 50, 75 or 100%.
- Category 1 rewards per review: 25% Rs100, 50% Rs150, 75% Rs250, 100% Rs500. Category 2: 25% Rs150, 50% Rs250, 75% Rs400, 100% Rs750.
- Earning = reward per review of the slab x total accepted reviews of that category (the technician total is the sum of both categories).
- Projection: reviews still needed to reach the next slab, and the earning if that slab is reached.

Sections of the website
- Slab rules are shown at the top of the page.
- Review dashboard: summary of review tickets uploaded, accepted (counted as reviews) and rejected (not counted); franchise ranking with search, sort and a review % filter; per franchise the technician cards with slab ladder, earnings and reviews needed for the next slab (search and slab filter).
- Dashboard based on new rules: the same calculations and franchise/technician views, limited to tickets uploaded in the selected report month and the preceding month. The original dashboard remains unchanged.
- Rejected tickets: review tickets that are not counted, with reason (call Cancelled, closed but other dealer, not in the calls file), franchise, technician and call details. Tabs: Summary, By franchise, All, and one per reason. Every table has a search bar and a filter arrow (▾) on each column (tick values, sort A-Z / Z-A, Clear filters).
- Rejection insights: key Chennai + ROTN and all-branch findings from Rating_Rejection_Analysis.xlsx, including the leading reasons and prevention checks.
- If all uploaded tickets are accepted: every uploaded review ticket, including the rejected ones, is accepted. The rejected tickets are added to the reviews of their franchise / technician / category (the closed-ticket denominator does not change). It shows review %, earnings and the extra money compared with now.
- Download Excel (top right): tick any subsections (for example Franchise summary, Technician slabs, Rejected - Cancelled, All accepted - Technicians) and download them as one workbook (one sheet per subsection) or as separate Excel files. The Excel always contains all rows; filters and search are not applied.
- Removed in this version: Order ID issues, Possibility of rejection and Earnings if no rejection (page sections, admin buttons and calculation code).

Notes
- Branches: the review Excel rows with Branch = Chennai or ROTN are used. The Calls Registered file must contain both branches (you can also select two calls files together, one per branch).
- The old ticket check and Order ID category section has been removed from the public page.
- The site is public. It shows franchise names and technician names. Order IDs are never stored on the site.
