import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# Create necessary directories
os.makedirs("data", exist_ok=True)
os.makedirs("reports", exist_ok=True)

class AttendanceAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Attendance Data Analyzer")
        self.root.geometry("500x500")

        # Heading
        ttk.Label(root, text="Student Attendance Analyzer", font=("Arial", 16, "bold")).pack(pady=10)

        # Upload Button
        ttk.Button(root, text="Upload Attendance CSV", command=self.load_and_analyze).pack(pady=10)

        # Date range selection
        ttk.Label(root, text="Start Date (YYYY-MM-DD)").pack(pady=5)
        self.start_date_entry = ttk.Entry(root)
        self.start_date_entry.pack(pady=5)

        ttk.Label(root, text="End Date (YYYY-MM-DD)").pack(pady=5)
        self.end_date_entry = ttk.Entry(root)
        self.end_date_entry.pack(pady=5)

        # Analyze Button
        ttk.Button(root, text="Analyze Attendance", command=self.analyze_with_date_filter).pack(pady=10)

        # Export Report Button
        ttk.Button(root, text="Export Attendance Report (PDF)", command=self.export_report_pdf).pack(pady=10)

    def load_and_analyze(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        try:
            self.df = pd.read_csv(file_path)
            self.df.to_csv(os.path.join("data", os.path.basename(file_path)), index=False)
            messagebox.showinfo("File Loaded", f"File '{os.path.basename(file_path)}' loaded successfully!")
            self.analyze_attendance(self.df)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def analyze_attendance(self, df):
        try:
            required_columns = {"Student Name", "Date", "Status"}
            if not required_columns.issubset(df.columns):
                raise ValueError("CSV must contain 'Student Name', 'Date', and 'Status' columns.")

            # Calculate attendance percentage
            attendance_summary = df.groupby("Student Name")["Status"].apply(lambda x: (x == "Present").mean() * 100)
            
            # Plot attendance distribution
            plt.figure(figsize=(10, 6))
            sns.barplot(x=attendance_summary.index, y=attendance_summary.values, palette="magma")
            plt.xticks(rotation=45, ha="right")
            plt.title("Student Attendance Percentage")
            plt.ylabel("Attendance (%)")
            plt.xlabel("Student Name")
            plt.tight_layout()
            plt.savefig("reports/attendance_summary.png")
            plt.show()

            # Display summary
            summary_message = "\n".join([f"{name}: {pct:.2f}%" for name, pct in attendance_summary.items()])
            messagebox.showinfo("Attendance Summary", summary_message)

            # Insights
            most_regular = attendance_summary.idxmax()
            least_regular = attendance_summary.idxmin()
            overall_attendance = attendance_summary.mean()

            insights = (
                f"🎯 Most Regular Student: {most_regular} ({attendance_summary[most_regular]:.2f}%)\n"
                f"🚶‍♂️ Least Regular Student: {least_regular} ({attendance_summary[least_regular]:.2f}%)\n"
                f"📊 Overall Class Attendance: {overall_attendance:.2f}%"
            )
            messagebox.showinfo("Attendance Insights", insights)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to analyze attendance: {e}")

    def analyze_with_date_filter(self):
        try:
            start_date = self.start_date_entry.get()
            end_date = self.end_date_entry.get()

            if not (start_date and end_date):
                raise ValueError("Please enter both start and end dates.")

            start_date = datetime.strptime(start_date, "%Y-%m-%d")
            end_date = datetime.strptime(end_date, "%Y-%m-%d")

            self.df['Date'] = pd.to_datetime(self.df['Date'])
            filtered_df = self.df[(self.df['Date'] >= start_date) & (self.df['Date'] <= end_date)]

            if filtered_df.empty:
                messagebox.showinfo("No Data", "No records found for the selected date range.")
                return

            self.analyze_attendance(filtered_df)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to filter data: {e}")

    def export_report_pdf(self):
        try:
            if not hasattr(self, 'df'):
                messagebox.showwarning("No Data", "Please upload and analyze attendance data first.")
                return

            pdf_file = "reports/attendance_report.pdf"
            doc = SimpleDocTemplate(pdf_file, pagesize=A4, title="Attendance Report")

            # Prepare data for the table
            table_data = [list(self.df.columns)] + self.df.values.tolist()

            # Create the table
            table = Table(table_data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))

            # Add table to PDF
            elements = [Paragraph("📊 Attendance Report", getSampleStyleSheet()["Title"]), table]
            doc.build(elements)

            messagebox.showinfo("Report Exported", f"PDF report saved to: {pdf_file}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report: {e}")

# Run the application
root = tk.Tk()
app = AttendanceAnalyzerApp(root)
root.mainloop()
