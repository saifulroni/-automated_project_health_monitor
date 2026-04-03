from dashboard_generator import generate_dashboard
from report_generator import build_pdf_report
from gantt_generator import generate_all_gantts
from email_report import send_email_report
from datetime import datetime
import os

os.makedirs('outputs', exist_ok=True)
date_str = datetime.today().strftime('%Y%m%d')

pdf_path = f'outputs/lemu-tajin_report_{date_str}.pdf'
dashboard_path = f'outputs/lemu-tajin_dashboard_{date_str}.xlsx'

print("=" * 50)
print("LEMU TAJIN LAUNCH — PMO HEALTH MONITOR")
print(f"Running: {datetime.today().strftime('%A, %d %B %Y')}")
print("=" * 50)

print("\n[1/4] Generating Gantt charts...")
generate_all_gantts('Tasks.xlsx')

print("\n[2/4] Generating Excel dashboard...")
generate_dashboard('Tasks.xlsx', dashboard_path)

print("\n[3/4] Generating PDF report...")
build_pdf_report('Tasks.xlsx', pdf_path)

print("\n[4/4] Sending email...")
send_email_report(
    sender_email='################',
    app_password='##################',
    recipient_email='#################',
    pdf_path=pdf_path
)

print("\n✓ All outputs saved to /outputs/")
print("=" * 50)