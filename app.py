# ==========================================
# 9. CLINICAL GENERATION & DOWNLOAD HUB
# ==========================================
st.divider()
st.markdown("### 3️⃣ Generate & Download Clinical Drafts (FBA & BIP)")

st.markdown(
    "Select your preferred report language below. The generated Word documents"
    " will include bilingual headings and content to support international"
    " clinical teams and bilingual family reviews."
)

lang_choice = st.radio(
    "Select Report Language Format:",
    options=[
        "English Only (纯英文报告)",
        "Bilingual: English + Chinese (双语：英文 + 中文)",
        "Bilingual: English + Spanish (Bilingüe: Inglés + Español)",
    ],
    index=0,
    horizontal=True,
)

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
  st.markdown("#### 📄 Functional Behavior Assessment (FBA)")
  st.info(
      "Generates a comprehensive 6-section FBA report including individual"
      " behavior tracking charts and QABF results."
  )

  # Multi-select for behaviors to include in FBA
  selected_behaviors_for_fba = []
  st.markdown("**Select Behaviors to Include in FBA Report:**")
  for idx, b in enumerate(active_behaviors):
    is_checked = st.checkbox(
        f"{b['name']}", value=True, key=f"fba_b_{selected_cohort_key}_{idx}"
    )
    if is_checked:
      selected_behaviors_for_fba.append(b)

  if len(selected_behaviors_for_fba) == 0:
    st.warning("Please select at least one behavior for the FBA report.")
  else:
    fba_docx_io = generate_exact_fba_doc(
        selected_cohort_key, lang_choice, selected_behaviors_for_fba
    )
    st.download_button(
        label=f"📥 Download FBA Draft (.docx)",
        data=fba_docx_io,
        file_name=(
            f"FBA_Draft_{current_meta['file_tag']}_"
            f"{'zh' if 'Chinese' in lang_choice else 'es' if 'Spanish' in lang_choice else 'en'}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

with col_dl2:
  st.markdown("#### 📑 Behavior Intervention Plan (BIP)")
  st.info(
      "Generates an enriched 9-section BIP document including crisis safety"
      " protocols, FCT, DRA, and staff training plans."
  )

  bip_docx_io = generate_exact_bip_doc(selected_cohort_key, lang_choice)
  st.download_button(
      label=f"📥 Download BIP Draft (.docx)",
      data=bip_docx_io,
      file_name=(
          f"BIP_Draft_{current_meta['file_tag']}_"
          f"{'zh' if 'Chinese' in lang_choice else 'es' if 'Spanish' in lang_choice else 'en'}.docx"
      ),
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )

st.divider()
st.markdown(
    "<div style='text-align: center; color: #7f8c8d; font-size: 0.85rem;'>"
    "BCBA Clinical FBA & BIP Draft Formulation Tool v2.6 | Designed for"
    " Professional Behavior Analysts | 100% Local Processing & HIPAA Compliant"
    " Architecture"
    "</div>",
    unsafe_allow_html=True,
)
```[cite: 12]
