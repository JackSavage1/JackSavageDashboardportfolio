"""
SOSi Linguist Management Dashboard
A comprehensive Streamlit dashboard for tracking linguist requests and availability
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np

# Page configuration
st.set_page_config(
    page_title="SOSi Linguist Management Dashboard",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .urgent-text {
        color: #ff4b4b;
        font-weight: bold;
    }
    .success-text {
        color: #00cc88;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Title and header
st.title("🌐 SOSi Linguist Management Dashboard")
st.markdown("---")

# Sidebar for file upload and filters
with st.sidebar:
    st.header("📁 Data Upload")
    
    # File uploaders
    master_file = st.file_uploader(
        "Upload Master Data (OutlookSOSitableScraping.xlsx)",
        type=['xlsx', 'xls']
    )
    
    analysis_file = st.file_uploader(
        "Upload Analysis Data (AllDataAnalysis.xlsx)",
        type=['xlsx', 'xls']
    )
    
    linguist_file = st.file_uploader(
        "Upload Linguist Data (SOSiApprovedLinguists.xlsx)",
        type=['xlsx', 'xls']
    )
    
    st.markdown("---")
    
    # Date range filter
    st.header("📅 Date Filters")
    date_range = st.date_input(
        "Select Date Range",
        value=(datetime.now() - timedelta(days=180), datetime.now()),
        max_value=datetime.now()
    )

# Load and process data function
@st.cache_data
def load_data(master_file, analysis_file, linguist_file):
    """Load and process uploaded Excel files"""
    data = {}
    
    if master_file:
        data['master'] = pd.read_excel(master_file)
        # Clean column names
        data['master'].columns = data['master'].columns.str.strip()
    
    if analysis_file:
        # Read the specific sheets
        try:
            data['analysis'] = pd.read_excel(analysis_file, sheet_name='All Historical Data')
            data['summary'] = pd.read_excel(analysis_file, sheet_name='Summary')
        except:
            data['analysis'] = pd.read_excel(analysis_file)
    
    if linguist_file:
        data['linguists'] = pd.read_excel(linguist_file)
    
    return data

# Process language statistics
def process_language_stats(df):
    """Process language request statistics"""
    if 'Language' not in df.columns:
        return pd.DataFrame()
    
    # Count requests per language
    language_counts = df['Language'].value_counts().reset_index()
    language_counts.columns = ['Language', 'Request Count']
    
    # Add sourcing status if available
    if 'Has Linguist?' in df.columns:
        sourcing_status = df.groupby('Language')['Has Linguist?'].apply(
            lambda x: 'Sourceable' if (x == 'YES').any() else 'Not Sourceable'
        ).reset_index()
        sourcing_status.columns = ['Language', 'Status']
        language_counts = language_counts.merge(sourcing_status, on='Language', how='left')
    
    return language_counts

# Main dashboard
if master_file or analysis_file:
    # Load data
    data = load_data(master_file, analysis_file, linguist_file)
    
    # Create tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Overview", 
        "🌍 Language Analysis", 
        "📈 Trends", 
        "⚠️ Gap Analysis",
        "📋 Detailed Reports"
    ])
    
    # Tab 1: Overview
    with tab1:
        st.header("Executive Summary")
        
        # Calculate key metrics
        if 'analysis' in data:
            df = data['analysis']
            
            # Key metrics in columns
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_requests = len(df)
                st.metric("Total Requests", f"{total_requests:,}")
            
            with col2:
                if 'Has Linguist?' in df.columns:
                    fulfilled = (df['Has Linguist?'] == 'YES').sum()
                    fulfillment_rate = (fulfilled / total_requests * 100) if total_requests > 0 else 0
                    st.metric("Fulfillment Rate", f"{fulfillment_rate:.1f}%")
            
            with col3:
                if 'Language' in df.columns:
                    unique_languages = df['Language'].nunique()
                    st.metric("Unique Languages", unique_languages)
            
            with col4:
                if 'Has Linguist?' in df.columns:
                    unfulfilled = (df['Has Linguist?'] == 'NO').sum()
                    st.metric("Unfulfilled Requests", unfulfilled, delta=f"-{unfulfilled}")
            
            st.markdown("---")
            
            # Top languages chart
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Top 10 Most Requested Languages")
                if 'Language' in df.columns:
                    top_languages = df['Language'].value_counts().head(10)
                    fig = px.bar(
                        x=top_languages.values,
                        y=top_languages.index,
                        orientation='h',
                        labels={'x': 'Number of Requests', 'y': 'Language'},
                        color_discrete_sequence=['#1f77b4']
                    )
                    fig.update_layout(height=400, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.subheader("Fulfillment Status Distribution")
                if 'Has Linguist?' in df.columns:
                    status_counts = df['Has Linguist?'].value_counts()
                    fig = px.pie(
                        values=status_counts.values,
                        names=['Sourceable' if x == 'YES' else 'Not Sourceable' for x in status_counts.index],
                        color_discrete_map={'Sourceable': '#00cc88', 'Not Sourceable': '#ff4b4b'}
                    )
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
    
    # Tab 2: Language Analysis
    with tab2:
        st.header("Language Request Analysis")
        
        if 'analysis' in data:
            df = data['analysis']
            language_stats = process_language_stats(df)
            
            # Filter options
            col1, col2 = st.columns([1, 3])
            with col1:
                status_filter = st.selectbox(
                    "Filter by Status",
                    ["All", "Sourceable", "Not Sourceable"]
                )
            
            # Apply filter
            if status_filter != "All" and 'Status' in language_stats.columns:
                language_stats = language_stats[language_stats['Status'] == status_filter]
            
            # Display sortable table
            st.subheader("Language Request Details")
            st.dataframe(
                language_stats,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Request Count": st.column_config.NumberColumn(
                        "Requests",
                        format="%d",
                    ),
                    "Status": st.column_config.TextColumn(
                        "Status",
                    ),
                }
            )
            
            # Language gap visualization
            st.subheader("Top Languages Without Linguists")
            if 'Has Linguist?' in df.columns:
                no_linguist_df = df[df['Has Linguist?'] == 'NO']
                if not no_linguist_df.empty:
                    gap_languages = no_linguist_df['Language'].value_counts().head(15)
                    fig = px.bar(
                        x=gap_languages.index,
                        y=gap_languages.values,
                        labels={'x': 'Language', 'y': 'Unfulfilled Requests'},
                        color_discrete_sequence=['#ff4b4b']
                    )
                    fig.update_layout(height=400, xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)
    
    # Tab 3: Trends
    with tab3:
        st.header("Request Trends Over Time")
        
        if 'master' in data:
            df = data['master']
            
            # Convert date columns
            date_columns = ['Date of request', 'Hearing Date', 'Row Added']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Time series analysis
            if 'Row Added' in df.columns:
                df_time = df.dropna(subset=['Row Added'])
                df_time['Month'] = df_time['Row Added'].dt.to_period('M').astype(str)
                monthly_requests = df_time.groupby('Month').size().reset_index(name='Requests')
                
                fig = px.line(
                    monthly_requests,
                    x='Month',
                    y='Requests',
                    title='Monthly Request Volume',
                    markers=True
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            
            # Day of week analysis
            col1, col2 = st.columns(2)
            
            with col1:
                if 'Row Added' in df.columns:
                    df_time['Day of Week'] = df_time['Row Added'].dt.day_name()
                    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                    day_counts = df_time['Day of Week'].value_counts().reindex(day_order, fill_value=0)
                    
                    fig = px.bar(
                        x=day_counts.index,
                        y=day_counts.values,
                        title='Requests by Day of Week',
                        labels={'x': 'Day', 'y': 'Number of Requests'}
                    )
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                if 'Priority' in df.columns:
                    priority_counts = df['Priority'].value_counts()
                    fig = px.pie(
                        values=priority_counts.values,
                        names=priority_counts.index,
                        title='Request Priority Distribution',
                        color_discrete_map={
                            'URGENT': '#ff4b4b',
                            'HIGH': '#ffa500',
                            'NORMAL': '#00cc88'
                        }
                    )
                    st.plotly_chart(fig, use_container_width=True)
    
    # Tab 4: Gap Analysis
    with tab4:
        st.header("Linguist Gap Analysis")
        
        if 'analysis' in data:
            df = data['analysis']
            
            # Critical gaps
            st.subheader("🚨 Critical Language Gaps")
            st.markdown("Languages with 5+ unfulfilled requests")
            
            if 'Has Linguist?' in df.columns:
                no_linguist_df = df[df['Has Linguist?'] == 'NO']
                if not no_linguist_df.empty:
                    gap_summary = no_linguist_df['Language'].value_counts()
                    critical_gaps = gap_summary[gap_summary >= 5].reset_index()
                    critical_gaps.columns = ['Language', 'Unfulfilled Requests']
                    critical_gaps['Priority'] = critical_gaps['Unfulfilled Requests'].apply(
                        lambda x: 'CRITICAL' if x >= 10 else 'HIGH' if x >= 5 else 'MEDIUM'
                    )
                    
                    # Style the dataframe
                    st.dataframe(
                        critical_gaps,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Priority": st.column_config.TextColumn(
                                "Priority",
                                help="Based on number of unfulfilled requests"
                            ),
                        }
                    )
            
            # Regional analysis if location data available
            if 'Location' in df.columns or 'Notes' in df.columns:
                st.subheader("📍 Geographic Distribution")
                location_col = 'Location' if 'Location' in df.columns else 'Notes'
                location_counts = df[location_col].value_counts().head(10)
                
                fig = px.bar(
                    x=location_counts.index,
                    y=location_counts.values,
                    title='Top 10 Request Locations',
                    labels={'x': 'Location', 'y': 'Number of Requests'}
                )
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
    
    # Tab 5: Detailed Reports
    with tab5:
        st.header("Detailed Reports")
        
        # Export options
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📥 Download Full Analysis"):
                if 'analysis' in data:
                    csv = data['analysis'].to_csv(index=False)
                    st.download_button(
                        label="Download CSV",
                        data=csv,
                        file_name=f"sosi_analysis_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
        
        with col2:
            if st.button("📊 Generate Executive Report"):
                st.info("Executive report generation in progress...")
        
        with col3:
            if st.button("📧 Email Report"):
                st.info("Email functionality to be implemented")
        
        # Detailed data view
        st.subheader("Raw Data View")
        
        data_selection = st.selectbox(
            "Select Dataset",
            ["Master Data", "Analysis Data", "Linguist Data"]
        )
        
        if data_selection == "Master Data" and 'master' in data:
            st.dataframe(data['master'], use_container_width=True, height=400)
        elif data_selection == "Analysis Data" and 'analysis' in data:
            st.dataframe(data['analysis'], use_container_width=True, height=400)
        elif data_selection == "Linguist Data" and 'linguists' in data:
            st.dataframe(data['linguists'], use_container_width=True, height=400)
        
        # Search functionality
        st.subheader("🔍 Search Records")
        search_term = st.text_input("Search by COI, Language, or Requester")
        
        if search_term and 'analysis' in data:
            df = data['analysis']
            # Search across multiple columns
            mask = df.apply(lambda x: x.astype(str).str.contains(search_term, case=False).any(), axis=1)
            filtered_df = df[mask]
            
            if not filtered_df.empty:
                st.write(f"Found {len(filtered_df)} matching records")
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning("No matching records found")

else:
    # Welcome screen when no data is loaded
    st.info("👈 Please upload your data files using the sidebar to begin analysis")
    
    # Instructions
    with st.expander("📖 How to Use This Dashboard"):
        st.markdown("""
        ### Getting Started
        
        1. **Upload Your Data Files**
           - OutlookSOSitableScraping.xlsx (Master hearing request data)
           - AllDataAnalysis.xlsx (Processed analysis data)
           - SOSiApprovedLinguists.xlsx (Available linguist data)
        
        2. **Navigate Through Tabs**
           - **Overview**: High-level metrics and summary
           - **Language Analysis**: Detailed language request breakdown
           - **Trends**: Time-based analysis and patterns
           - **Gap Analysis**: Critical linguist gaps and priorities
           - **Detailed Reports**: Raw data and export options
        
        3. **Use Filters**
           - Date range selector in sidebar
           - Status filters in Language Analysis tab
           - Search functionality in Detailed Reports
        
        ### Key Metrics Explained
        
        - **Fulfillment Rate**: Percentage of requests with available linguists
        - **Critical Gaps**: Languages with 5+ unfulfilled requests
        - **Priority Levels**: 
          - CRITICAL: 10+ unfulfilled requests
          - HIGH: 5-9 unfulfilled requests
          - MEDIUM: 3-4 unfulfilled requests
        """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center'>
        <p>SOSi Linguist Management Dashboard v1.0 | Last Updated: {}</p>
    </div>
    """.format(datetime.now().strftime("%Y-%m-%d %H:%M")),
    unsafe_allow_html=True
)