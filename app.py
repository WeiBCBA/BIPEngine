# 🌟 UPDATED CRITICAL DE-IDENTIFICATION & DEMO NOTICE CARD
st.markdown(
    """
    <div class="critical-security-card">
        <div class="security-title">
            🛡️ DATA PRIVACY & DE-IDENTIFICATION PROTOCOL
        </div>
        <div class="security-body">
            <ul>
                <li><strong>Current Live Demo Notice:</strong> This interactive portal is designed strictly for demonstration purposes. All downloadable sample datasets provided on this page are 100% synthetic, standardized, and fully <strong>de-identified</strong>. You can safely test, upload, and evaluate the engine with complete peace of mind.</li>
                <li><strong>Future Local Deployment Workflow:</strong> In future production environments deployed directly on the clinician's local workstation, the engine will automatically execute an end-to-end <strong>local de-identification pipeline</strong> before parsing any raw uploaded files.</li>
                <li><strong>Seamless Anonymization:</strong> All generated FBA & BIP drafts replace sensitive Personal Identifiable Information (PII) with structured placeholders (e.g., <code>[CLIENT_NAME]</code>, <code>[FACILITY_NAME]</code>).</li>
                <li><strong>Finalization:</strong> Clinicians simply press <strong>CTRL + H</strong> (Find & Replace) in Microsoft Word to insert real client identifiers and perform clinical edits prior to signature. Absolute data privacy and compliance are guaranteed across all phases!</li>
            </ul>
        </div>
    </div>
""",
    unsafe_allow_html=True,
)
