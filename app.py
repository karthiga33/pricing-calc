import streamlit as st
import pandas as pd
import tempfile
import os
import requests
from datetime import datetime
from test2 import CostReportAgent

@st.cache_data(ttl=86400)
def fetch_usd_to_inr():
    try:
        resp = requests.get("https://open.er-api.com/v6/latest/USD", timeout=5)
        data = resp.json()
        if data.get("result") == "success":
            return data["rates"]["INR"]
    except Exception:
        pass
    return 85.50

st.set_page_config(page_title="AWS Cost Estimator (Nova Pro)", page_icon="☁️", layout="wide")

st.markdown("""
<style>
    .stApp {background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);}
    .main .block-container {padding: 2rem 3rem; background: rgba(255, 255, 255, 0.95); border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); margin-top: 2rem;}
    h1 {color: #2c3e50; font-weight: 800; text-align: center; font-size: 3rem !important; margin-bottom: 0.5rem;}
    .subtitle {text-align: center; color: #666; font-size: 1.2rem; margin-bottom: 2rem; font-weight: 300;}
    .stTextInput label, .stNumberInput label, .stFileUploader label {color: #333; font-weight: 600; font-size: 1rem;}
    .stFileUploader {background: linear-gradient(135deg, #e8f4f8 0%, #d4e7f1 100%); padding: 1.5rem; border-radius: 15px; border: 2px dashed #3498db;}
    .stButton > button {background: linear-gradient(135deg, #3498db 0%, #2980b9 100%); color: white; border: none; padding: 0.75rem 2rem; font-size: 1.1rem; font-weight: 600; border-radius: 50px; box-shadow: 0 4px 15px rgba(52, 152, 219, 0.3); width: 100%;}
    .stButton > button:hover {transform: translateY(-2px); box-shadow: 0 6px 20px rgba(52, 152, 219, 0.5);}
    .stDownloadButton > button {background: linear-gradient(135deg, #27ae60 0%, #229954 100%); color: white; border: none; padding: 0.75rem 2rem; font-size: 1.1rem; font-weight: 600; border-radius: 50px; box-shadow: 0 4px 15px rgba(39, 174, 96, 0.3); width: 100%;}
    .stSuccess {background: linear-gradient(135deg, #d5f4e6 0%, #c8e6c9 100%); border-left: 5px solid #27ae60; border-radius: 10px; padding: 1rem;}
    .stError {background: linear-gradient(135deg, #fadbd8 0%, #f5b7b1 100%); color: #c0392b; border-radius: 10px; padding: 1rem;}
    .stTextInput > div > div > input, .stNumberInput > div > div > input {border-radius: 10px; border: 2px solid #e0e0e0; padding: 0.75rem;}
    .stTextInput > div > div > input:focus, .stNumberInput > div > div > input:focus {border-color: #3498db; box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);}
    hr {margin: 2rem 0; border: none; height: 2px; background: linear-gradient(90deg, transparent, #3498db, transparent);}
</style>
""", unsafe_allow_html=True)

st.markdown("# ☁️ AWS Cost Estimator (Nova Pro)")
st.markdown('<p class="subtitle">Generate comprehensive AWS cost reports with AWS Bedrock Nova Pro</p>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📋 About")
    st.info("""
    This tool generates detailed AWS cost estimation reports including:
    
    ✅ EC2 instance specifications  
    ✅ Cost breakdowns (USD & INR)  
    ✅ AI-powered service descriptions  
    ✅ Best practices recommendations  
    """)
    
    st.markdown("### 📊 CSV Requirements")
    st.warning("""
    Your CSV must contain:
    - **Service** column
    - **Monthly Cost** column
    - **Configuration Summary** column
    
    *First 7 rows will be skipped*
    """)
    
    st.markdown("### 🔧 Powered By")
    st.markdown("""
    - **AWS Bedrock Nova Pro**
    - Streamlit
    - OpenPyXL
    """)
    
    st.markdown("### ⚙️ Setup Instructions")
    st.info("""
    **AWS Bedrock Setup:**
    1. Configure AWS credentials
    2. Run `aws configure`
    3. Ensure Bedrock access in us-east-1
    """)

col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📁 Upload CSV File")
    uploaded_file = st.file_uploader("Choose your AWS cost CSV file", type=["csv"], label_visibility="collapsed")
    
    st.markdown("### 👤 Customer Information")
    customer_name = st.text_input("Customer Name", placeholder="Enter customer name...")
    region = st.text_input("AWS Region", value="US East (N. Virginia)")

live_rate = fetch_usd_to_inr()
default_output = customer_name.strip().replace(' ', '_') if customer_name.strip() else ""

with col2:
    st.markdown("### 💰 Pricing Configuration")
    st.info(f"📡 Live USD → INR rate: ₹{live_rate:.2f} (updated daily)")
    usd_to_inr = st.number_input("USD to INR Exchange Rate", min_value=0.0, value=live_rate, step=0.01, format="%.2f")
    pricing_link = st.text_input("Pricing Link (Optional)", placeholder="https://calculator.aws/...")
    output_filename = st.text_input("Output File Name", value=default_output, placeholder="e.g., AWS_Cost_Report")

st.markdown("---")

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    generate_btn = st.button("🚀 Generate Cost Report", use_container_width=True)

if generate_btn:
    if not uploaded_file:
        st.error("❌ Please upload a CSV file")
    elif not customer_name.strip():
        st.error("❌ Please enter a customer name")
    elif not output_filename.strip():
        st.error("❌ Please enter an output file name")
    else:
        final_filename = output_filename.strip()
        if not final_filename.lower().endswith('.xlsx'):
            final_filename += '.xlsx'
        
        with st.spinner("🔄 Generating your cost report with Nova Pro... This may take a moment..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp_file:
                    tmp_file.write(uploaded_file.read())
                    tmp_file_path = tmp_file.name
                
                output_file_path = os.path.join(tempfile.gettempdir(), final_filename)
                
                agent = CostReportAgent(default_usd_to_inr=usd_to_inr, default_region=region)
                result = agent.generate_cost_report(
                    input_file=tmp_file_path,
                    output_file=output_file_path,
                    customer_name=customer_name.strip(),
                    usd_to_inr=usd_to_inr,
                    region=region,
                    pricing_link=pricing_link.strip()
                )
                
                try:
                    os.remove(tmp_file_path)
                except:
                    pass
                
                if result["status"] == "success":
                    st.success("✅ Cost report generated successfully!")
                    
                    st.markdown("### 📊 Report Preview")
                    
                    try:
                        df_summary = pd.read_excel(result["file"], sheet_name="Summary", dtype=str)
                        df_services = pd.read_excel(result["file"], sheet_name="AWS Services", dtype=str)
                        
                        tab1, tab2 = st.tabs(["💰 Cost Summary", "🔧 AWS Services"])
                        
                        with tab1:
                            st.dataframe(df_summary, use_container_width=True)
                        
                        with tab2:
                            st.dataframe(df_services, use_container_width=True)
                    except Exception as e:
                        st.warning(f"⚠️ Preview unavailable: {e}")
                    
                    st.markdown("---")
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        with open(result["file"], "rb") as f:
                            st.download_button(
                                label="📥 Download Excel Report",
                                data=f,
                                file_name=final_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                    
                    try:
                        os.remove(output_file_path)
                    except:
                        pass
                        
                else:
                    st.error(f"❌ Error: {result['message']}")
                    
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.info("💡 Make sure AWS credentials are configured with Bedrock access")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 2rem 0;'>
    <p>Made with ❤️ using AWS Bedrock Nova Pro | © 2024</p>
</div>
""", unsafe_allow_html=True)
