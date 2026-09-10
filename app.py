import streamlit as st
import pandas as pd
import plotly.express as px
import os,glob,re

st.set_page_config(page_title="Rural Business Advisor",page_icon="🌾",layout="wide")
BASE=os.path.dirname(os.path.abspath(__file__))

# ---------- EXCEL DATA ----------
def clean_sheet(path,sheet,expected=None):
    raw=pd.read_excel(path,sheet_name=sheet,header=None)
    header=0
    if expected:
        for i,row in raw.iterrows():
            vals=[str(x).strip() for x in row.tolist()]
            if expected in vals:
                header=i
                break
    df=pd.read_excel(path,sheet_name=sheet,header=header)
    df=df.dropna(how="all")
    df.columns=[str(c).strip() for c in df.columns]
    return df

xlsx=glob.glob(os.path.join(BASE,"*.xlsx"))
market_file=None
business_file=None
scheme_file=None
for f in xlsx:
    try:
        sheets=pd.ExcelFile(f).sheet_names
        if "Sheet2" in sheets and "Sheet3" in sheets and "Sheet4" in sheets:
            market_file=f
        elif len(sheets)==1:
            cols=[str(c).strip().lower() for c in pd.read_excel(f).columns]
            if "business" in cols and any("starting cost" in c for c in cols):
                business_file=f
        if "scheme" in os.path.basename(f).lower() or "loan" in os.path.basename(f).lower():
            scheme_file=f
    except:
        pass

# Business cost/demand workbook
if business_file:
    try:
        braw=pd.read_excel(business_file)
        braw.columns=[str(c).strip() for c in braw.columns]
        def money_num(x):
            nums=re.findall(r'\d+(?:\.\d+)?',str(x).replace(",",""))
            if not nums:return 0
            vals=[float(n) for n in nums]
            return vals[0]*100000 if "lakh" in str(x).lower() else vals[0]
        bdf=pd.DataFrame({
            "Business":braw.get("Business",pd.Series(dtype=str)).astype(str).str.strip(),
            "Starting Cost":braw.get("Starting Cost",pd.Series(dtype=str)).apply(money_num),
            "Demand Text":braw.get("Demand",pd.Series(dtype=str)).astype(str),
            "Competition Text":braw.get("Competition",pd.Series(dtype=str)).astype(str),
            "Profit Text":braw.get("Profit",pd.Series(dtype=str)).astype(str),
            "Risk":braw.get("Main Risks",pd.Series(dtype=str)).astype(str)
        })
        def score_demand(x):
            s=str(x).lower()
            if "very high" in s:return 10
            if "high" in s:return 9
            if "medium" in s:return 7
            if "low" in s:return 5
            return 7
        def score_comp(x):
            s=str(x).lower()
            if "very high" in s:return 10
            if "high" in s:return 9
            if "medium-high" in s or "medium high" in s:return 8
            if "medium" in s:return 6
            if "low" in s:return 4
            return 6
        bdf["Demand Score"]=bdf["Demand Text"].apply(score_demand)
        bdf["Competition Score"]=bdf["Competition Text"].apply(score_comp)
        # Keep profit text from Excel and estimate numeric profit only when a clear number exists.
        bdf["Expected Profit"]=bdf["Profit Text"].apply(lambda x: float(re.findall(r'\d+(?:\.\d+)?',str(x).replace(",",""))[0]) if re.findall(r'\d+(?:\.\d+)?',str(x).replace(",","")) else 0)
        df=bdf
    except Exception as e:
        df=pd.DataFrame()
        business_error=str(e)
else:
    df=pd.DataFrame()
    business_error="Business Excel file not found."

# Market workbook: Sheet1 Business, Sheet2 Market, Sheet3 Competitor, Sheet4 Local
market_df=competitor_df=local_df=pd.DataFrame()
market_error=None
if market_file:
    try:
        market_df=clean_sheet(market_file,"Sheet2","Market Segment")
        competitor_df=clean_sheet(market_file,"Sheet3","Competitor")
        local_df=clean_sheet(market_file,"Area")
    except Exception as e:
        market_error=str(e)
else:
    market_error="Market Excel file not found."

# Government schemes: only use a workbook whose name indicates schemes/loan.
schemes_df=pd.DataFrame()
scheme_error="Government scheme Excel not found."
if scheme_file:
    try:
        schemes_df=pd.read_excel(scheme_file)
        schemes_df.columns=[str(c).strip() for c in schemes_df.columns]
        scheme_error=None
    except Exception as e:
        scheme_error=str(e)

# ---------- FUNCTIONS ----------
def calculate_feasibility(business,budget,land,water,experience):
    row=df[df["Business"]==business].iloc[0]
    score=row["Demand Score"]*4+(10-row["Competition Score"])*3
    if budget>=row["Starting Cost"]:score+=20
    elif budget>=row["Starting Cost"]*.7:score+=12
    else:score+=5
    if land=="Yes":score+=5
    if water=="Yes":score+=5
    if experience=="Yes":score+=5
    return min(int(score),100)

def get_risk(score):
    return "Low" if score>=75 else "Medium" if score>=55 else "High"

def calculate_emi(principal,rate,years):
    if principal<=0 or years<=0:return 0
    if rate==0:return principal/(years*12)
    r=rate/12/100;n=years*12
    return principal*r*(1+r)**n/((1+r)**n-1)

# ---------- SIDEBAR ----------
st.sidebar.title("🌾 Rural Business Advisor")
option=st.sidebar.selectbox("Choose a page",[
    "🏠 Home","💼 Business Recommendation","📊 Market Analysis",
    "💰 Financial Calculator","🏦 Scheme Recommendation","📄 Business Report","ℹ️ About"
])

if option=="🏠 Home":
    st.title("🌾 Rural Business Advisor")
    st.subheader("AI-assisted hyper-local business advisory and financial planning")
    st.write("This platform helps rural entrepreneurs understand suitable business opportunities, market conditions, financial requirements and possible government financing options.")
    if not df.empty and not market_df.empty:
        st.success(f"✅ Live project data loaded: {len(df)} businesses, {len(market_df)} market records, {len(competitor_df)} competitors and {len(local_df)} local records.")
    else:
        st.warning("Some project Excel data could not be loaded. Check that the Excel files are in the same folder as app.py.")
    st.header("🎯 Our Objective")
    st.write("Help rural entrepreneurs make better business decisions using location, available capital, demand, competition, profitability, financial planning and government schemes.")
    st.header("⭐ Key Features")
    c1,c2,c3=st.columns(3)
    with c1: st.subheader("💡 Business Recommendation");st.write("Evaluate business suitability based on budget, resources and market factors.")
    with c2: st.subheader("📍 Hyper-Local Analysis");st.write("Analyze local market, population, footfall and competitors.")
    with c3: st.subheader("📊 Market Analysis");st.write("Compare demand, competition, investment and expected profit.")
    c1,c2,c3=st.columns(3)
    with c1: st.subheader("💰 Financial Planning");st.write("Calculate investment, loan requirement, profit, EMI and payback period.")
    with c2: st.subheader("🏦 Scheme Routing");st.write("Use government loan scheme information from the Finance team's dataset.")
    with c3: st.subheader("📄 Business Report");st.write("Combine business, market and financial information into one report.")
    st.header("⚙️ How It Works")
    st.write("**1. Enter Location & Requirements** → **2. Analyze Business Opportunity** → **3. Calculate Feasibility** → **4. Prepare Financial Plan** → **5. Find Government Schemes** → **6. Generate Business Report**")
    st.success("🌾 One platform for business feasibility, financial planning and scheme discovery.")

elif option=="💼 Business Recommendation":
    st.title("💼 Business Recommendation")
    if df.empty:
        st.error("Business Excel data is not loaded.");st.stop()
    location=st.text_input("📍 Enter your location",placeholder="Example: Barasat")
    budget=st.number_input("💰 Available Capital (₹)",min_value=0,step=5000)
    land=st.selectbox("🌱 Agricultural Land Available?",["Yes","No"])
    water=st.selectbox("💧 Reliable Water Supply?",["Yes","No"])
    experience=st.selectbox("👨‍🌾 Experience in Farming/Business?",["Yes","No"])
    interest=st.selectbox("🏪 Business Interest",df["Business"].tolist())
    market_radius=st.selectbox("📍 Target Market Reach",["5 km","10 km"])
    if st.button("🔍 Analyze Business"):
        if not location:st.warning("Please enter your location.")
        elif budget==0:st.warning("Please enter your available capital.")
        else:
            row=df[df["Business"]==interest].iloc[0];score=calculate_feasibility(interest,budget,land,water,experience);risk=get_risk(score)
            st.success("Business analysis completed!")
           st.subheader("🎯 Business Assessed");st.success(interest);st.metric("⭐ Feasibility Score",f"{score}/100");st.progress(score/100)
            c1,c2,c3,c4=st.columns(4)
            c1.metric("📈 Demand",f"{row['Demand Score']}/10");c2.metric("🏪 Competition",f"{row['Competition Score']}/10");c3.metric("💰 Starting Cost",f"₹{row['Starting Cost']:,.0f}");c4.metric("⚠️ Risk",risk)
            st.subheader("📌 Data from Excel")
            c1,c2=st.columns(2)
            with c1:
                st.write(f"**Demand:** {row['Demand Text']}")
                st.write(f"**Competition:** {row['Competition Text']}")
            with c2:
                st.write(f"**Expected Profit:** {row['Profit Text']}")
                st.write(f"**Main Risks:** {row['Risk']}")
            st.subheader("📍 Hyper-Local Market")
            st.write(f"Target market reach: **{market_radius}** around **{location}**.")
            if not local_df.empty:
                st.dataframe(local_df,use_container_width=True,hide_index=True)
            st.subheader("💡 Why This Business?")
            reasons=[]
            if budget>=row["Starting Cost"]:reasons.append("✓ Available capital covers the estimated starting cost.")
            if row["Demand Score"]>=8:reasons.append("✓ The Excel dataset shows strong demand.")
            if row["Competition Score"]<=6:reasons.append("✓ Competition is relatively manageable.")
            if experience=="Yes":reasons.append("✓ Previous experience can support implementation.")
            if not reasons:reasons.append("✓ Business selected based on the entered requirements.")
            for r in reasons:st.write(r)
            st.subheader("📋 SWOT Analysis")
            c1,c2=st.columns(2)
            with c1:
                st.markdown("### 💪 Strengths");st.write("• Local demand opportunity\n• Potential for small-scale operation\n• Can use local resources")
                st.markdown("### 🚀 Opportunities");st.write("• Expansion to nearby markets\n• Direct/local selling\n• Possible financing support")
            with c2:
                st.markdown("### ⚠️ Weaknesses");st.write("• Initial investment requirement\n• Need to validate local costs")
                st.markdown("### 🔴 Threats");st.write("• Price fluctuations\n• Seasonal demand\n• Supply chain problems\n• Local competition")

elif option=="📊 Market Analysis":
    st.title("📊 Hyper-Local Market Analysis")
    if df.empty:st.error("Business Excel data is not loaded.");st.stop()
    location=st.text_input("📍 Enter Location",placeholder="Example: Barasat")
    radius=st.selectbox("📍 Market Reach",["5 km","10 km"])
    selected_business=st.selectbox("🏪 Select Business",df["Business"].tolist())
    selected=df[df["Business"]==selected_business].iloc[0]
    st.subheader("📍 Selected Business")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Demand",f"{selected['Demand Score']}/10");c2.metric("Competition",f"{selected['Competition Score']}/10");c3.metric("Starting Cost",f"₹{selected['Starting Cost']:,.0f}");c4.metric("Risk",get_risk(calculate_feasibility(selected_business,max(selected["Starting Cost"],1),"No","No","No")))
    if not market_df.empty:
        st.subheader("📈 Market Data from Excel")
        st.dataframe(market_df,use_container_width=True,hide_index=True)
    if not competitor_df.empty:
        st.subheader("🏪 Competitor Data from Excel")
        st.dataframe(competitor_df,use_container_width=True,hide_index=True)
    if not local_df.empty:
        st.subheader("📍 Local Data from Excel")
        st.dataframe(local_df,use_container_width=True,hide_index=True)
        if "Population" in local_df.columns:
            st.write("The Local Data sheet contains population, customer groups, business types, footfall, local demand and business opportunity information.")
    st.subheader("📊 Business Comparison")
    st.plotly_chart(px.bar(df,x="Business",y="Demand Score",title="Demand Score from Business Excel"),use_container_width=True)
    st.plotly_chart(px.bar(df,x="Business",y="Competition Score",title="Competition Score from Business Excel"),use_container_width=True)
    st.subheader("📋 Business Data")
    display_cols=["Business","Starting Cost","Demand Text","Competition Text","Profit Text","Risk"]
    st.dataframe(df[display_cols],use_container_width=True,hide_index=True)
    st.info(f"📍 Analysis area: {location if location else 'Location not entered'} | Market reach: {radius}")

elif option=="💰 Financial Calculator":
    st.title("💰 Smart Financial Calculator")
    st.write("Estimate project cost, financing requirement, monthly profit, EMI and payback period.")
    c1,c2=st.columns(2)
    with c1:
        project_cost=st.number_input("💰 Total Project Cost (₹)",min_value=0,step=5000)
        own_capital=st.number_input("💵 Own Contribution / Margin Capital (₹)",min_value=0,step=5000)
        monthly_income=st.number_input("📈 Expected Monthly Income (₹)",min_value=0,step=1000)
    with c2:
        monthly_expenses=st.number_input("💸 Monthly Expenses (₹)",min_value=0,step=1000)
        interest_rate=st.number_input("📊 Annual Interest Rate (%)",min_value=0.0,value=8.0,step=.5)
        repayment_years=st.number_input("📅 Repayment Period (Years)",min_value=1,max_value=20,value=5)
   
    if st.button("🧮 Calculate Financial Plan"):
        if project_cost==0:st.warning("Please enter project cost.")
        elif own_capital>project_cost:st.warning("Own contribution cannot be greater than project cost.")
        else:
            loan_required=project_cost-own_capital;monthly_profit=monthly_income-monthly_expenses;emi=calculate_emi(loan_required,interest_rate,repayment_years);payback_period = own_capital / monthly_profit if monthly_profit > 0 else 0
            st.success("Financial calculation completed!")
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Project Cost",f"₹{project_cost:,.0f}");c2.metric("Own Contribution",f"₹{own_capital:,.0f}");c3.metric("Loan Required",f"₹{loan_required:,.0f}");c4.metric("Monthly Profit",f"₹{monthly_profit:,.0f}")
            c1,c2=st.columns(2)
            c1.metric("Estimated EMI",f"₹{emi:,.0f}")
            c2.metric("Interest Rate",f"{interest_rate}%")
            if monthly_profit>0:st.success(f"Estimated payback period on own contribution: {payback_period:.1f} months")
            else:st.error("Monthly expenses are equal to or higher than income.")
            months=repayment_years*12;remaining=float(loan_required);schedule=[];monthly_rate=interest_rate/12/100
            for month in range(1,months+1):
                if remaining<=0:break
                interest=0 if interest_rate==0 else remaining*monthly_rate
                principal=min(emi,remaining) if interest_rate==0 else max(0,min(emi-interest,remaining))
                remaining=max(0,remaining-principal)
                schedule.append({"Month":month,"EMI":round(emi,2),"Interest":round(interest,2),"Principal":round(principal,2),"Remaining Loan":round(remaining,2)})
            st.subheader("📅 Repayment Schedule");st.dataframe(pd.DataFrame(schedule),use_container_width=True,hide_index=True)
            st.info("📌 EMI is an estimate. Actual lender terms may differ.")

elif option=="🏦 Scheme Recommendation":
    st.title("🏦 Government Scheme Recommendation")
    st.write("Find financing schemes based on your business and project requirements.")
    business=st.selectbox("🏪 Business Type",["Agriculture","Dairy","Poultry","Fish Farming","Food Processing","Tailoring","Grocery Shop"])
    applicant=st.selectbox("👤 Applicant Type",["Individual","Farmer","FPO/SHG/Cooperative","Street Vendor","Micro/Small Enterprise"])
    gender=st.selectbox("👩 Woman Entrepreneur?",["Yes","No"]);scst=st.selectbox("📋 SC/ST Applicant?",["Yes","No"]);status=st.selectbox("📋 Business Status",["New Business","Existing Business"])
    investment=st.number_input("💰 Available Own Capital (₹)",min_value=0,step=5000);project_cost=st.number_input("💰 Proposed Project Cost (₹)",min_value=0,step=5000)
    infrastructure=st.selectbox("🏗️ Is this an agriculture/post-harvest infrastructure project?",["No","Yes"])
    if st.button("🔍 Find Suitable Schemes"):
        if project_cost==0:st.warning("Please enter the proposed project cost.")
        elif investment>project_cost:st.warning("Own capital cannot be greater than project cost.")
        else:
            loan=max(0,project_cost-investment);recommended=[]
            if applicant=="Street Vendor" and loan<=50000:recommended.append(("PM SVANidhi","Suitable for eligible street vendors with financing up to ₹50,000."))
            if applicant=="Farmer" and business in ["Agriculture","Dairy","Poultry","Fish Farming"]:recommended.append(("Kisan Credit Card (KCC)","Potentially suitable for eligible farmers and agricultural/allied activities."))
            if infrastructure=="Yes" and loan<=20000000:recommended.append(("Agriculture Infrastructure Fund (AIF)","Potentially suitable for eligible agriculture and post-harvest infrastructure projects."))
            if (gender=="Yes" or scst=="Yes") and status=="New Business" and 1000000<=loan<=10000000:recommended.append(("Stand-Up India","Potentially suitable for eligible women or SC/ST entrepreneurs starting a new greenfield enterprise."))
            if loan<=2000000:recommended.append(("Pradhan Mantri MUDRA Yojana (PMMY)","Suitable for eligible micro businesses within the MUDRA limit."))
            if status=="New Business":
                if business=="Food Processing" and project_cost<=5000000:recommended.append(("PMEGP","Potentially suitable for a new eligible manufacturing micro enterprise."))
                elif business in ["Tailoring","Grocery Shop"] and project_cost<=2000000:recommended.append(("PMEGP","Potentially suitable for a new eligible service/business micro enterprise."))
            if applicant=="Micro/Small Enterprise" and loan<=100000000:recommended.append(("Credit Guarantee Scheme through CGTMSE","May provide credit guarantee support for eligible micro/small enterprises; it is not a direct loan."))
            st.subheader("🎯 Recommended Schemes")
            if recommended:
                for name,reason in recommended:st.success(f"🏦 {name}");st.write(f"**Why:** {reason}")
            else:st.info("No clear scheme match found from the entered information.")
            st.subheader("📋 Scheme Details")
            if not schemes_df.empty:st.dataframe(schemes_df,use_container_width=True,hide_index=True)
            else:st.info("Government scheme Excel is not connected in this version.")
            st.warning("⚠️ This is a preliminary rule-based recommendation. Final eligibility, interest rates and loan terms must be verified from official scheme guidelines.")

elif option=="📄 Business Report":
    st.title("📄 Business Report")
    if df.empty:st.error("Business Excel data is not loaded.");st.stop()
    location=st.text_input("📍 Location",placeholder="Example: Barasat");budget=st.number_input("💰 Available Capital (₹)",min_value=0,step=5000)
    business=st.selectbox("🏪 Business",df["Business"].tolist());land=st.selectbox("🌱 Land Available?",["Yes","No"]);water=st.selectbox("💧 Water Available?",["Yes","No"]);experience=st.selectbox("👨‍🌾 Experience?",["Yes","No"]);market_radius=st.selectbox("📍 Market Reach",["5 km","10 km"])
    if st.button("📄 Generate Report"):
        if not location or budget==0:st.warning("Please enter location and available capital.")
        else:
            row=df[df["Business"]==business].iloc[0];score=calculate_feasibility(business,budget,land,water,experience);risk=get_risk(score)
            st.success("Report generated!");st.header("🌾 Rural Business Feasibility Report")
            st.write(f"**Location:** {location}");st.write(f"**Business:** {business}");st.write(f"**Market Reach:** {market_radius}")
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Feasibility",f"{score}/100");c2.metric("Starting Cost",f"₹{row['Starting Cost']:,.0f}");c3.metric("Expected Profit",row["Profit Text"]);c4.metric("Risk",risk)
            st.subheader("📊 Business Data from Excel")
            st.write(f"**Demand:** {row['Demand Text']}");st.write(f"**Competition:** {row['Competition Text']}");st.write(f"**Main Risks:** {row['Risk']}")
            st.subheader("💡 Recommendation")
            if score>=75:st.success("This business shows strong preliminary feasibility based on the entered information.")
            elif score>=55:st.warning("This business shows moderate feasibility. Check local demand, costs and competition before investing.")
            else:st.error("This business has lower preliminary feasibility. Consider comparing other business options.")
            st.info("⚠️ This report is a prototype decision-support output and should not be treated as guaranteed financial advice.")

else:
    st.title("ℹ️ About the Project")
    st.subheader("🌾 Rural Business Advisor")
    st.write("Rural Business Advisor is a decision-support platform designed to help rural entrepreneurs evaluate business opportunities, understand local market conditions, plan finances and explore government financing options.")
    st.header("❓ Problem");st.write("Rural entrepreneurs may face difficulty understanding which business is suitable for their locality, how much investment is required and which financing options may be relevant.")
    st.header("💡 Solution");st.write("Our platform combines hyper-local business analysis, business recommendation, financial calculation and government scheme information in one platform.")
    st.header("🛠️ Technologies");st.write("🐍 Python  |  🌐 Streamlit  |  🐼 Pandas  |  📊 Plotly ")
    st.header("🚀 Future Development");st.write("• Machine-learning based business recommendation\n• Real competitor mapping\n• Multilingual and voice support\n• More accurate government scheme routing\n• Digital Market place Integration")
    st.success("🌾 Helping rural entrepreneurs make smarter business decisions.")
