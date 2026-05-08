import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# --- 1. CONFIG & STYLE ---
st.set_page_config(page_title="4P's Engagement Portal", layout="wide")

# Danh sách Admin có quyền xem toàn bộ công ty
SUPER_ADMIN_IDS = ["PZ016155", "PZ007411", "PZ004485", "PZEX0011", "PZEX0001"]
DRIVER_MAPPING = {
    'Vision & Purpose': 'Q2', 'Growth & Autonomy': 'Q3', 
    'Team & Culture': 'Q4', 'Well-being': 'Q5', 'Appreciation': 'Q6'
}

# --- 2. GLOBAL UTILS ---
def standardize_id(x):
    val = str(x).strip().upper()
    if val in ['NAN', '0', '', 'NONE']: return 'NONE'
    if val.isdigit(): return f"PZ{val.zfill(6)}"
    return val

def apply_4ps_color(val):
    if val == "Hidden": 
        return 'color: #95a5a6; font-style: italic;'
    try:
        val = float(val)
        if val < 3.5: return 'background-color: #e74c3c; color: black; font-weight: bold;'
        if val < 4.0: return 'background-color: #f1c40f; color: black; font-weight: bold;'
        return 'background-color: #2ecc71; color: black; font-weight: bold;'
    except: 
        return ''

def safe_format(val):
    try:
        return f"{float(val):.2f}"
    except (ValueError, TypeError):
        return val

# --- 3. DATA LOADING ---
@st.cache_data
def load_all_data():
    f_emp, f_rep = 'Employee list.xlsx', 'Engagement Report - April cycle.xlsx'
    if not os.path.exists(f_emp) or not os.path.exists(f_rep): return [None]*8

    df_emp = pd.read_excel(f_emp)
    df_emp.columns = [str(c).strip() for c in df_emp.columns]
    mgr_col = 'Manager Code' if 'Manager Code' in df_emp.columns else 'Manager'
    df_emp['Employee Code'] = df_emp['Employee Code'].apply(standardize_id)
    df_emp['Manager_Clean'] = df_emp[mgr_col].apply(standardize_id)

    df_s1 = pd.read_excel(f_rep, sheet_name='Sheet1')
    df_t1 = pd.read_excel(f_rep, sheet_name='Table1')
    
    def clean_cols(df):
        df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
        return df

    df_s1, df_t1 = clean_cols(df_s1), clean_cols(df_t1)
    df_s1['ID'] = df_s1['ID'].apply(standardize_id)
    df_t1['Subject ID'] = df_t1.iloc[:, 0].apply(standardize_id)

    def find_col(df, keywords):
        for col in df.columns:
            if all(k.lower() in col.lower() for k in keywords): return col
        return df.columns[0]

    H_S1 = [find_col(df_s1, ["group", "1"]), find_col(df_s1, ["department", "2"]), 
            find_col(df_s1, ["function", "3"]), find_col(df_s1, ["team", "4"])]
    H_T1 = [find_col(df_t1, ["group", "1"]), find_col(df_t1, ["department", "2"]), 
            find_col(df_t1, ["function", "3"]), find_col(df_t1, ["team", "4"])]

    Q8_cols = [c for c in df_t1.columns if 'Q8' in c]
    Q9_cols = [c for c in df_t1.columns if 'Q9' in c]
    Q8, Q9 = Q8_cols[0], Q9_cols[0]
    
    drivers = list(DRIVER_MAPPING.keys())
    for name, prefix in DRIVER_MAPPING.items():
        cols = [c for c in df_t1.columns if c.startswith(prefix)]
        df_t1[name] = pd.to_numeric(df_t1[cols].mean(axis=1), errors='coerce')
    df_t1['Average Score'] = df_t1[drivers].mean(axis=1)

    return df_emp, df_s1, df_t1, Q8, Q9, H_S1, H_T1, drivers

data_bundle = load_all_data()
if data_bundle[0] is None: 
    st.error("Missing Excel files! Please check filenames on GitHub.")
    st.stop()

df_emp, data_s1, data_t1, Q8, Q9, H_S1, H_T1, DRIVER_NAMES = data_bundle

# --- 4. RECURSIVE LOOKUP (SƠ ĐỒ QUẢN LÝ) ---
def get_all_subordinates(mgr_id, emp_df):
    subs = emp_df[emp_df['Manager_Clean'] == mgr_id]['Employee Code'].unique().tolist()
    all_subs = []
    stack = subs.copy()
    while stack:
        curr = stack.pop()
        if curr not in all_subs:
            all_subs.append(curr)
            children = emp_df[emp_df['Manager_Clean'] == curr]['Employee Code'].unique().tolist()
            stack.extend(children)
    return all_subs

# --- 5. APP INTERFACE ---
with st.sidebar:
    if os.path.exists("4pslogo.png"): 
        st.image("4pslogo.png", width=150)
    u_id_raw = st.text_input("🔑 Employee ID:").strip().upper()

if u_id_raw:
    u_id = standardize_id(u_id_raw)
    is_admin = u_id in SUPER_ADMIN_IDS
    
    if is_admin or u_id in df_emp['Employee Code'].values:
        # Thông tin cá nhân sếp để làm Benchmark
        my_info = data_s1[data_s1['ID'] == u_id]
        my_parent_group = my_info[H_S1[0]].values[0] if not my_info.empty else None
        my_parent_dept = my_info[H_S1[1]].values[0] if not my_info.empty else None
        my_parent_func = my_info[H_S1[2]].values[0] if not my_info.empty else None

        # Xác định phạm vi dữ liệu
        if is_admin:
            full_scope_ids = data_s1['ID'].unique().tolist()
        else:
            full_scope_ids = [u_id] + get_all_subordinates(u_id, df_emp)
        
        # Chỉ lấy những người thực tế đã làm khảo sát
        responded_ids = data_t1[data_t1['Subject ID'].isin(full_scope_ids)]['Subject ID'].unique().tolist()
        
        scope_s1 = data_s1[data_s1['ID'].isin(full_scope_ids)]
        scope_t1 = data_t1[data_t1['Subject ID'].isin(responded_ids)]
        
        with st.sidebar:
            def get_opts(df, c): return sorted([str(x) for x in df[c].dropna().unique() if str(x) not in ['0', 'nan']])
            
            f1 = st.selectbox("Group", ["All"] + get_opts(scope_s1, H_S1[0]))
            d2_s1 = scope_s1 if f1=="All" else scope_s1[scope_s1[H_S1[0]]==f1]
            d2_t1 = scope_t1 if f1=="All" else scope_t1[scope_t1[H_T1[0]]==f1]
            
            f2 = st.selectbox("Department", ["All"] + get_opts(d2_s1, H_S1[1]))
            d3_s1 = d2_s1 if f2=="All" else d2_s1[d2_s1[H_S1[1]]==f2]
            d3_t1 = d2_t1 if f2=="All" else d2_t1[d2_t1[H_T1[1]]==f2]
            
            f3 = st.selectbox("Function", ["All"] + get_opts(d3_s1, H_S1[2]))
            d4_s1 = d3_s1 if f3=="All" else d3_s1[d3_s1[H_S1[2]]==f3]
            d4_t1 = d3_t1 if f3=="All" else d3_t1[d3_t1[H_T1[2]]==f3]
            
            f4 = st.selectbox("Team", ["All"] + get_opts(d4_s1, H_S1[3]))
            df_f_s1 = d4_s1 if f4=="All" else d4_s1[d4_s1[H_S1[3]]==f4]
            df_f_t1 = d4_t1 if f4=="All" else d4_t1[d4_t1[H_T1[3]]==f4]

        st.title("🍕 Engagement Dashboard")
        
        # Row 1: Metrics
        total_staff = len(df_f_s1)
        responses = len(df_f_t1)
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Staff", total_staff)
        m2.metric("Responses Received", responses)
        m3.metric("Response Rate", f"{(responses/total_staff*100):.1f}%" if total_staff > 0 else "0%")

        # KIỂM TRA BẢO MẬT (THRESHOLD < 5)
        if responses < 5:
            st.warning("⚠️ **Privacy Shield:** Data is hidden because the number of responses is less than 5.")
        else:
            st.divider()
            st.subheader("🎯 Employee NPS (eNPS)")
            nps_res = pd.to_numeric(df_f_t1[Q8], errors='coerce').dropna()
            if not nps_res.empty:
                total_n = len(nps_res)
                p = len(nps_res[nps_res>=9])
                pa = len(nps_res[(nps_res>=7)&(nps_res<=8)])
                d = len(nps_res[nps_res<=6])
                enps_val = ((p-d)/total_n)*100
                
                c1, c2 = st.columns([1, 1.5])
                c1.metric("eNPS Score", f"{enps_val:.1f}")
                
                fig_nps_pie = px.pie(
                    values=[p, pa, d], 
                    names=[f'Promoters: {(p/total_n*100):.1f}%', f'Passives: {(pa/total_n*100):.1f}%', f'Detractors: {(d/total_n*100):.1f}%'],
                    color_discrete_sequence=['#2ecc71', '#f1c40f', '#e74c3c'], 
                    hole=0.4
                )
                fig_nps_pie.update_traces(sort=False) 
                c2.plotly_chart(fig_nps_pie.update_layout(height=350, margin=dict(t=30,b=30)), use_container_width=True)

        # Row 3: Hierarchy Comparison
        st.divider()
        st.subheader("📊 Hierarchy Score Comparison")
        
        def get_sum(df, label, force_show=False):
            if df.empty: return None
            actual_resp = len(df)
            if actual_resp < 5 and not force_show:
                return {'Level': f"{label} (Hidden)", 'Avg Score': "Hidden", **{dr: "Hidden" for dr in DRIVER_NAMES}}
            return {'Level': label, 'Avg Score': df['Average Score'].mean(), **{dr: df[dr].mean() for dr in DRIVER_NAMES}}
        
        rows = []
        rows.append(get_sum(data_t1, "🌍 TOTAL 4P'S (Benchmark)", force_show=True))
        if not is_admin:
            if my_parent_group: rows.append(get_sum(data_t1[data_t1[H_T1[0]] == my_parent_group], f"🏢 Group: {my_parent_group}", force_show=True))
            if my_parent_dept: rows.append(get_sum(data_t1[data_t1[H_T1[1]] == my_parent_dept], f"📁 Dept: {my_parent_dept}", force_show=True))
            if my_parent_func: rows.append(get_sum(data_t1[data_t1[H_T1[2]] == my_parent_func], f"📍 Store: {my_parent_func}", force_show=True))
        
        rows.append(get_sum(scope_t1, "⭐ YOUR TEAM TOTAL"))
        
        if f1 == "All":
            for g in get_opts(scope_t1, H_T1[0]): rows.append(get_sum(scope_t1[scope_t1[H_T1[0]]==g], f"  ↳ Group: {g}"))
        elif f2 == "All":
            for d in get_opts(d2_t1, H_T1[1]): rows.append(get_sum(d2_t1[d2_t1[H_T1[1]]==d], f"  ↳ Dept: {d}"))
        elif f3 == "All":
            for f in get_opts(d3_t1, H_T1[2]): rows.append(get_sum(d3_t1[d3_t1[H_T1[2]]==f], f"  ↳ Func: {f}"))
        else:
            for t in get_opts(d4_t1, H_T1[3]): rows.append(get_sum(d4_t1[d4_t1[H_T1[3]]==t], f"  ↳ Team: {t}"))

        res_df = pd.DataFrame([r for r in rows if r]).set_index('Level')
        
        # Hiển thị bảng với format an toàn
        st.table(res_df.style.format(safe_format).map(apply_4ps_color))

        # Row 4: Feedback Analysis (Chỉ hiện khi đủ 5 phản hồi)
        if responses >= 5:
            st.divider()
            st.subheader("💬 Feedback Analysis")
            df_f_t1['NPS_Group'] = df_f_t1[Q8].apply(lambda x: "Promoters" if x>=9 else ("Passives" if x>=7 else "Detractors"))
            
            for grp in ["Promoters", "Passives", "Detractors"]:
                with st.expander(f"🔍 Analysis: {grp}"):
                    g_df = df_f_t1[df_f_t1['NPS_Group']==grp]
                    if not g_df.empty:
                        b_data = []
                        for dr in DRIVER_NAMES:
                            b_data.append({'Factor': dr, 'Count': len(g_df[g_df[dr]>=4.0]), 'Type': 'Continue'})
                            b_data.append({'Factor': dr, 'Count': -len(g_df[g_df[dr]<3.5]), 'Type': 'Improve'})
                        
                        st.plotly_chart(px.bar(pd.DataFrame(b_data), x='Count', y='Factor', color='Type', 
                                               orientation='h', 
                                               color_discrete_map={'Continue':'#2ecc71','Improve':'#e74c3c'}), 
                                        use_container_width=True)
                        
                        for c in g_df[Q9].dropna(): 
                            if str(c).strip(): st.info(f"“{c}”")
    else:
        st.sidebar.error("Access Denied. ID not found or unauthorized.")
