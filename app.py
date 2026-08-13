with tab2:
    st.subheader(f"🤝 Stakeholder Input ({subj_en}-Centered)")
    
    # ------------------ 新增：允许上传 Stakeholder 访谈/反馈文件 ------------------
    st.markdown("##### 📂 Upload Stakeholder Notes / Interviews (Optional)")
    stakeholder_file = st.file_uploader(
        "Upload Teacher/Parent Interview notes or questionnaire (.txt, .docx, .csv):", 
        type=["txt", "docx", "csv"], 
        key="stakeholder_file_uploader"
    )
    
    uploaded_stakeholder_text = ""
    if stakeholder_file is not None:
        try:
            if stakeholder_file.name.endswith(".txt"):
                uploaded_stakeholder_text = stakeholder_file.read().decode("utf-8")
            elif stakeholder_file.name.endswith(".docx"):
                doc_obj = docx.Document(stakeholder_file)
                uploaded_stakeholder_text = "\n".join([p.text for p in doc_obj.paragraphs if p.text.strip()])
            elif stakeholder_file.name.endswith(".csv"):
                df_stk = pd.read_csv(stakeholder_file)
                uploaded_stakeholder_text = df_stk.to_string()
            st.success(f"Successfully read stakeholder file: '{stakeholder_file.name}'!")
        except Exception as e:
            st.error(f"Error reading file: {e}")
    # --------------------------------------------------------------------------------

    col_a, col_b = st.columns(2)
    with col_a:
        agency_name = st.text_input("School / Agency Name", "Metropolitan Inclusive Center")
        district_name = st.text_input("District / Health Region", "District 10 Behavioral Division")
        dob_val = st.text_input(f"{subj_en} DOB", "05/12/2022" if "Early" in selected_age_group else ("05/12/2015" if not is_adult else "05/12/2001"))
        id_val = st.text_input(f"{subj_en} ID", "ID-908231")
        fba_date = st.text_input("Date of FBA / BIP", "08/08/2026")
        
    with col_b:
        sources_options = ["Teacher Interview", "Parent Interview", "Rating Scales"]
        data_sources = st.multiselect("1. Secondary Data Sources (Direct Observations auto-included from Tab 1):", sources_options, default=sources_options)
        
    st.divider()
    st.markdown("### 2. Target Behavior Operational Breakdown & Examples")
    
    # 如果上传了文件，默认把文件提取的内容拼接到“行为描述”中；没上传就用默认文字
    default_desc = "Screaming (>80dB), pushing materials, dropping to floor during transitions or when demands are presented."
    if uploaded_stakeholder_text:
        default_desc = f"[Extracted from Stakeholder File]\n{uploaded_stakeholder_text[:300]}..."

    c1, c2, c3 = st.columns(3)
    target_beh = c1.text_area("Target Behavior Description", default_desc, height=100)
    beh_examples = c2.text_area("Examples of Target Behavior", "Throwing workbooks, yelling 'No!', hitting table with open palms.", height=100)
    beh_non_examples = c3.text_area("Non-Examples of Target Behavior", "Requesting 'Break' using PECS/AAC card, quietly sitting, asking for teacher assistance.", height=100)

    st.divider()
    st.markdown("### 3. Triggers & Behavioral Context")
    c_t1, c_t2 = st.columns(2)
    setting_events = c_t1.text_area("Setting Events (Slow Triggers)", "Overtiredness, lack of sleep, physical discomfort, or schedule changes.", height=70)
    antecedents_val = c_t2.text_area("Antecedent Events (Immediate Triggers)", "Presentation of multi-step academic tasks or transition away from preferred items.", height=70)
