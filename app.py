# 9. Hypothesis Statement
    doc.add_heading("9. Functional Behavioral Hypothesis Statement (行为功能假设说明)" if "Chinese" in selected_language else "9. Functional Behavioral Hypothesis Statement", level=1)
    primary_func = curr_defaults["primary_func"]

    clean_slow = setting_events.rstrip('.')
    clean_fast = antecedents_val.rstrip('.')
    clean_target = target_beh.rstrip('.')
    clean_cons = consequences_val.rstrip('.')

    hyp_en = (
        f"Under setting conditions of {clean_slow}, when presented with {clean_fast}, "
        f"[CLIENT_NAME] is likely to engage in {clean_target}. "
        f"This behavior is maintained by {clean_cons}, allowing the client to achieve {primary_func}. "
        f"The behavior serves as a functional communication attempt."
    )

    p9 = doc.add_paragraph()
    p9.add_run("Primary Hypothesis: ").bold = True
    p9.add_run(hyp_en)
    
    if "Chinese" in selected_language:
        # 确保所有变量都经过字典映射翻译
        zh_slow = translate(clean_slow, selected_language)
        zh_fast = translate(clean_fast, selected_language)
        zh_target = translate(clean_target, selected_language)
        zh_cons = translate(clean_cons, selected_language)
        zh_func = translate(primary_func, selected_language)

        hyp_zh = (
            f"在【{zh_slow}】的背景情境下，当出现【{zh_fast}】时，[CLIENT_NAME] 倾向于表现出【{zh_target}】。"
            f"该行为通过【{zh_cons}】得到维持，其核心功能在于【{zh_func}】。"
            f"该行为是传递沟通意图的功能性尝试。"
        )
        p9.add_run(f"\n（中文对照 - 行为假设表达: {hyp_zh}）").italic = True
