"""
Interactive Loan Scenario Analysis App
Clean and simple Streamlit app for analyzing mortgage refinancing scenarios
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import math
from datetime import datetime
import io

# Configure page
st.set_page_config(
    page_title="Loan Analyzer Pro",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

class LoanAnalyzer:
    def __init__(self, principal, annual_rate, years):
        self.principal = principal
        self.annual_rate = annual_rate
        self.years = years
        self.monthly_rate = annual_rate / 100 / 12
        self.num_payments = int(years * 12)
        
    def calculate_monthly_payment(self):
        if self.annual_rate == 0:
            return self.principal / self.num_payments
        
        payment = self.principal * (self.monthly_rate * (1 + self.monthly_rate)**self.num_payments) / ((1 + self.monthly_rate)**self.num_payments - 1)
        return payment
    
    def create_amortization_schedule(self, start_date=None):
        if start_date is None:
            start_date = datetime.now()
            
        monthly_payment = self.calculate_monthly_payment()
        schedule = []
        balance = self.principal
        cumulative_interest = 0
        current_date = start_date
        
        for payment_num in range(1, self.num_payments + 1):
            interest_payment = balance * self.monthly_rate
            principal_payment = monthly_payment - interest_payment
            balance -= principal_payment
            cumulative_interest += interest_payment
            
            schedule.append({
                'Payment_Number': payment_num,
                'Date': current_date.strftime('%Y-%m-%d'),
                'Beginning_Balance': balance + principal_payment,
                'Monthly_Payment': monthly_payment,
                'Principal_Payment': principal_payment,
                'Interest_Payment': interest_payment,
                'Ending_Balance': max(0, balance),
                'Cumulative_Interest': cumulative_interest
            })
            
            # Move to next month
            if current_date.month == 12:
                current_date = current_date.replace(year=current_date.year + 1, month=1)
            else:
                current_date = current_date.replace(month=current_date.month + 1)
                
            if balance <= 0:
                break
                
        return pd.DataFrame(schedule)

def create_amortization_chart(schedule_df, title="Amortization Schedule"):
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=('Principal vs Interest Over Time', 'Remaining Balance'),
        vertical_spacing=0.12
    )
    
    # Principal vs Interest
    fig.add_trace(
        go.Scatter(
            x=schedule_df['Payment_Number'],
            y=schedule_df['Principal_Payment'],
            name='Principal Payment',
            fill='tonexty',
            line=dict(color='#27AE60')
        ),
        row=1, col=1
    )
    
    fig.add_trace(
        go.Scatter(
            x=schedule_df['Payment_Number'],
            y=schedule_df['Interest_Payment'],
            name='Interest Payment',
            fill='tozeroy',
            line=dict(color='#E74C3C')
        ),
        row=1, col=1
    )
    
    # Remaining Balance
    fig.add_trace(
        go.Scatter(
            x=schedule_df['Payment_Number'],
            y=schedule_df['Ending_Balance'],
            name='Remaining Balance',
            line=dict(color='#3498DB', width=3)
        ),
        row=2, col=1
    )
    
    fig.update_layout(
        title=title,
        height=600,
        showlegend=True
    )
    
    return fig

def main():
    st.title("🏠 Loan Analyzer Pro")
    st.markdown("**Comprehensive mortgage refinancing analysis with independent loan control**")
    
    # Sidebar for loan inputs
    st.sidebar.header("🔧 Current Loan Parameters")
    
    # Apartment loan inputs
    st.sidebar.subheader("🏢 Apartment Loan")
    apt_principal = st.sidebar.number_input("Loan Amount ($)", value=555000, step=1000, key="apt_principal")
    apt_rate = st.sidebar.number_input("Interest Rate (%)", value=5.49, step=0.01, key="apt_rate")
    apt_term = st.sidebar.number_input("Loan Term (years)", value=27, step=1, key="apt_term")
    
    # Investment loan inputs
    st.sidebar.subheader("🏘️ Investment Property")
    inv_principal = st.sidebar.number_input("Loan Amount ($)", value=220000, step=1000, key="inv_principal")
    inv_rate = st.sidebar.number_input("Interest Rate (%)", value=6.49, step=0.01, key="inv_rate")
    inv_term = st.sidebar.number_input("Loan Term (years)", value=27.5, step=0.5, key="inv_term")
    
    # Create analyzers
    apartment_analyzer = LoanAnalyzer(apt_principal, apt_rate, apt_term)
    investment_analyzer = LoanAnalyzer(inv_principal, inv_rate, inv_term)
    
    # Main dashboard
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Dashboard", "🔄 Refinancing", "📈 Amortization", "📋 Schedules", "💾 Export"
    ])
    
    with tab1:
        st.header("Portfolio Overview")
        
        # Current loan summary
        col1, col2, col3, col4 = st.columns(4)
        
        apt_payment = apartment_analyzer.calculate_monthly_payment()
        inv_payment = investment_analyzer.calculate_monthly_payment()
        total_payment = apt_payment + inv_payment
        total_principal = apt_principal + inv_principal
        
        with col1:
            st.metric("Total Loan Amount", f"${total_principal:,.0f}")
        with col2:
            st.metric("Total Monthly Payment", f"${total_payment:,.0f}")
        with col3:
            weighted_rate = (apt_principal * apt_rate + inv_principal * inv_rate) / total_principal
            st.metric("Weighted Avg Rate", f"{weighted_rate:.2f}%")
        with col4:
            total_interest = (apt_payment * apt_term * 12 - apt_principal) + (inv_payment * inv_term * 12 - inv_principal)
            st.metric("Total Interest (Life of Loans)", f"${total_interest:,.0f}")
    
    with tab2:
        st.header("🔄 Independent Refinancing Analysis")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🏢 Apartment Loan")
            
            apt_refi = st.checkbox("Refinance Apartment Loan", key="apt_refi")
            
            if apt_refi:
                apt_new_rate = st.number_input("New Apartment Rate (%)", value=4.99, step=0.01, key="apt_new_rate")
                apt_new_term = st.number_input("New Apartment Term (years)", value=25, step=1, key="apt_new_term")
                apt_refi_costs = st.number_input("Apartment Refi Costs ($)", value=3000, step=100, key="apt_refi_costs")
                
                # Calculate new apartment payment
                new_apt_analyzer = LoanAnalyzer(apt_principal, apt_new_rate, apt_new_term)
                new_apt_payment = new_apt_analyzer.calculate_monthly_payment()
                apt_savings = apt_payment - new_apt_payment
                
                st.write(f"**Current Payment:** ${apt_payment:,.0f}")
                st.write(f"**New Payment:** ${new_apt_payment:,.0f}")
                
                if apt_savings > 0:
                    st.success(f"**Monthly Savings:** ${apt_savings:,.0f}")
                    st.success(f"**Break-even:** {apt_refi_costs / apt_savings:.1f} months")
                else:
                    st.error(f"**Monthly Increase:** ${abs(apt_savings):,.0f}")
            else:
                new_apt_payment = apt_payment
                apt_savings = 0
                apt_refi_costs = 0
        
        with col2:
            st.subheader("🏘️ Investment Property")
            
            inv_refi = st.checkbox("Refinance Investment Property", key="inv_refi")
            
            if inv_refi:
                inv_new_rate = st.number_input("New Investment Rate (%)", value=5.99, step=0.01, key="inv_new_rate")
                inv_new_term = st.number_input("New Investment Term (years)", value=25, step=1, key="inv_new_term")
                inv_refi_costs = st.number_input("Investment Refi Costs ($)", value=2000, step=100, key="inv_refi_costs")
                
                # Calculate new investment payment
                new_inv_analyzer = LoanAnalyzer(inv_principal, inv_new_rate, inv_new_term)
                new_inv_payment = new_inv_analyzer.calculate_monthly_payment()
                inv_savings = inv_payment - new_inv_payment
                
                st.write(f"**Current Payment:** ${inv_payment:,.0f}")
                st.write(f"**New Payment:** ${new_inv_payment:,.0f}")
                
                if inv_savings > 0:
                    st.success(f"**Monthly Savings:** ${inv_savings:,.0f}")
                    st.success(f"**Break-even:** {inv_refi_costs / inv_savings:.1f} months")
                else:
                    st.error(f"**Monthly Increase:** ${abs(inv_savings):,.0f}")
            else:
                new_inv_payment = inv_payment
                inv_savings = 0
                inv_refi_costs = 0
        
        # Combined results
        st.subheader("🏠 Combined Portfolio Results")
        
        total_new_payment = new_apt_payment + new_inv_payment
        total_savings = apt_savings + inv_savings
        total_refi_costs = apt_refi_costs + inv_refi_costs
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Current Total Payment", f"${total_payment:,.0f}")
        with col2:
            st.metric("New Total Payment", f"${total_new_payment:,.0f}")
        with col3:
            if total_savings > 0:
                st.metric("Total Monthly Savings", f"${total_savings:,.0f}", delta=f"+${total_savings:,.0f}")
            else:
                st.metric("Total Monthly Change", f"${total_savings:,.0f}", delta=f"{total_savings:,.0f}")
        with col4:
            if total_savings > 0 and total_refi_costs > 0:
                st.metric("Combined Break-even", f"{total_refi_costs / total_savings:.1f} months")
            else:
                st.metric("Combined Break-even", "N/A")
    
    with tab3:
        st.header("📈 Amortization Schedules")
        
        loan_choice = st.selectbox("Select Loan to Visualize", ["Apartment", "Investment Property", "Both"])
        
        if loan_choice in ["Apartment", "Both"]:
            st.subheader("🏢 Apartment Loan Amortization")
            
            # Choose current or refinanced
            if apt_refi:
                apt_schedule_choice = st.radio("Apartment Schedule", ["Current", "Refinanced"], key="apt_sched_choice")
                if apt_schedule_choice == "Refinanced":
                    display_analyzer = LoanAnalyzer(apt_principal, apt_new_rate, apt_new_term)
                else:
                    display_analyzer = apartment_analyzer
            else:
                display_analyzer = apartment_analyzer
            
            apt_schedule = display_analyzer.create_amortization_schedule()
            apt_chart = create_amortization_chart(apt_schedule, "Apartment Loan Amortization")
            st.plotly_chart(apt_chart, use_container_width=True)
        
        if loan_choice in ["Investment Property", "Both"]:
            st.subheader("🏘️ Investment Property Amortization")
            
            # Choose current or refinanced
            if inv_refi:
                inv_schedule_choice = st.radio("Investment Schedule", ["Current", "Refinanced"], key="inv_sched_choice")
                if inv_schedule_choice == "Refinanced":
                    display_analyzer = LoanAnalyzer(inv_principal, inv_new_rate, inv_new_term)
                else:
                    display_analyzer = investment_analyzer
            else:
                display_analyzer = investment_analyzer
            
            inv_schedule = display_analyzer.create_amortization_schedule()
            inv_chart = create_amortization_chart(inv_schedule, "Investment Property Amortization")
            st.plotly_chart(inv_chart, use_container_width=True)
    
    with tab4:
        st.header("📋 Detailed Payment Schedules")
        
        schedule_choice = st.selectbox("Select Schedule", [
            "Current Apartment",
            "Current Investment",
            "Refinanced Apartment" if apt_refi else None,
            "Refinanced Investment" if inv_refi else None
        ])
        
        if schedule_choice == "Current Apartment":
            schedule_df = apartment_analyzer.create_amortization_schedule()
        elif schedule_choice == "Current Investment":
            schedule_df = investment_analyzer.create_amortization_schedule()
        elif schedule_choice == "Refinanced Apartment" and apt_refi:
            refi_analyzer = LoanAnalyzer(apt_principal, apt_new_rate, apt_new_term)
            schedule_df = refi_analyzer.create_amortization_schedule()
        elif schedule_choice == "Refinanced Investment" and inv_refi:
            refi_analyzer = LoanAnalyzer(inv_principal, inv_new_rate, inv_new_term)
            schedule_df = refi_analyzer.create_amortization_schedule()
        else:
            st.warning("Please select a valid schedule option.")
            return
        
        # Display options
        show_full = st.checkbox("Show Full Schedule", value=False)
        
        if show_full:
            st.dataframe(schedule_df, use_container_width=True)
        else:
            st.subheader("First 12 Months")
            st.dataframe(schedule_df.head(12), use_container_width=True)
            
            yearly_data = schedule_df[schedule_df['Payment_Number'] % 12 == 0].head(10)
            if len(yearly_data) > 0:
                st.subheader("Year-End Summaries")
                st.dataframe(yearly_data, use_container_width=True)
    
    with tab5:
        st.header("💾 Export Data")
        
        export_options = [
            "Current Apartment Schedule",
            "Current Investment Schedule"
        ]
        
        if apt_refi:
            export_options.append("Refinanced Apartment Schedule")
        if inv_refi:
            export_options.append("Refinanced Investment Schedule")
        
        export_options.append("Portfolio Summary")
        
        export_choice = st.selectbox("Choose Data to Export", export_options)
        
        # Generate export data
        if export_choice == "Current Apartment Schedule":
            export_df = apartment_analyzer.create_amortization_schedule()
        elif export_choice == "Current Investment Schedule":
            export_df = investment_analyzer.create_amortization_schedule()
        elif export_choice == "Refinanced Apartment Schedule":
            refi_analyzer = LoanAnalyzer(apt_principal, apt_new_rate, apt_new_term)
            export_df = refi_analyzer.create_amortization_schedule()
        elif export_choice == "Refinanced Investment Schedule":
            refi_analyzer = LoanAnalyzer(inv_principal, inv_new_rate, inv_new_term)
            export_df = refi_analyzer.create_amortization_schedule()
        else:  # Portfolio Summary
            summary_data = {
                'Loan': ['Apartment', 'Investment'],
                'Principal': [apt_principal, inv_principal],
                'Current_Rate': [apt_rate, inv_rate],
                'Current_Term': [apt_term, inv_term],
                'Current_Payment': [apt_payment, inv_payment],
                'Refinanced': [apt_refi, inv_refi],
                'New_Rate': [apt_new_rate if apt_refi else apt_rate, inv_new_rate if inv_refi else inv_rate],
                'New_Term': [apt_new_term if apt_refi else apt_term, inv_new_term if inv_refi else inv_term],
                'New_Payment': [new_apt_payment, new_inv_payment],
                'Monthly_Savings': [apt_savings, inv_savings],
                'Refi_Costs': [apt_refi_costs, inv_refi_costs]
            }
            export_df = pd.DataFrame(summary_data)
        
        # Preview
        st.subheader("Data Preview")
        st.dataframe(export_df.head(10), use_container_width=True)
        
        # Download button
        csv_buffer = io.StringIO()
        export_df.to_csv(csv_buffer, index=False)
        csv_string = csv_buffer.getvalue()
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{export_choice.lower().replace(' ', '_')}_{timestamp}.csv"
        
        st.download_button(
            label=f"📥 Download {export_choice}",
            data=csv_string,
            file_name=filename,
            mime='text/csv'
        )

if __name__ == "__main__":
    main()