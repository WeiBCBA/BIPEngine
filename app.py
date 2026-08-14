from typing import Dict, List, Optional
import json

class FBABIPEngine:
    def __init__(self, language: str = "zh"):
        self.language = language
        self.supported_languages = ["zh", "en", "es"]
        
    def resolve_function_discrepancy(self, qabf_scores: Dict[str, float], abc_data: Dict[str, int]) -> Dict:
        """
        解决 QABF (间接问卷) 与 ABC (直接观察) 的数据冲突逻辑
        权重分配：ABC Data (0.65) > QABF Data (0.35)
        """
        # 归一化 ABC 频次为比例
        total_abc = sum(abc_data.values()) if sum(abc_data.values()) > 0 else 1
        abc_norm = {k: v / total_abc for k, v in abc_data.items()}
        
        # 归一化 QABF 得分为比例
        total_qabf = sum(qabf_scores.values()) if sum(qabf_scores.values()) > 0 else 1
        qabf_norm = {k: v / total_qabf for k, v in qabf_scores.items()}
        
        # 综合计算加权得分
        weighted_scores = {}
        all_functions = set(list(qabf_scores.keys()) + list(abc_data.keys()))
        
        for func in all_functions:
            score_abc = abc_norm.get(func, 0.0)
            score_qabf = qabf_norm.get(func, 0.0)
            weighted_scores[func] = round((score_abc * 0.65) + (score_qabf * 0.35), 2)
            
        # 排序寻找主要功能
        sorted_functions = sorted(weighted_scores.items(), key=lambda x: x[1], reverse=True)
        primary_function = sorted_functions[0][0]
        
        # 冲突检测：若 QABF 最高项与 ABC 最高项不一致，标记冲突警报
        qabf_top = max(qabf_norm, key=qabf_norm.get)
        abc_top = max(abc_norm, key=abc_norm.get)
        
        discrepancy_flag = False
        recommendation_note = ""
        
        if qabf_top != abc_top:
            discrepancy_flag = True
            recommendation_note = (
                f"警告：QABF 结果倾向于 [{qabf_top}]，而直接观察 ABC 数据强烈指向 [{abc_top}]。"
                "判定主导功能以 ABC 数据为主，建议进行短期 Functional Probes（功能探针测试）进一步验证。"
            )
            
        return {
            "primary_function": primary_function,
            "weighted_scores": weighted_scores,
            "discrepancy_flag": discrepancy_flag,
            "clinical_note": recommendation_note
        }

    def generate_fba_section_5(self, student_info: Dict, target_behavior: str, function_analysis: Dict) -> Dict:
        """
        生成第五点：个性化的 Synthesis and Clinical Recommendations
        """
        func = function_analysis["primary_function"]
        
        # 个性化 Hypothesis 总结
        hypothesis = (
            f"当处于 {student_info.get('antecedent_trigger', '特定触发情境')} 时，"
            f"{student_info['name']} 倾向于出现 {target_behavior}，"
            f"其主要功能在于获得/逃避 [{func}]。"
        )
        
        # 个性化环境调控与安全建议
        env_modifications = [
            f"在前因阶段提供预先视觉提示（Visual Schedule/First-Then Board），降低过渡焦虑。",
            f"针对 {func} 功能，提前安排 Scheduled Breaks 或 Sensory Satiation（感官饱食）。"
        ]
        
        safety_priorities = [
            "针对 Face Slapping / SIB 行为，确保环境无硬物，并在高发时段由一对一助教保持臂长距离关注。",
            "若发生连续自伤，立即执行 Block & Redirection（阻断与重定向），避免直接给予过度的语言关注。"
        ]
        
        return {
            "hypothesis_summary": hypothesis,
            "environmental_modifications": env_modifications,
            "safety_priorities": safety_priorities,
            "discrepancy_warning": function_analysis["clinical_note"]
        }

    def generate_bip_plan(self, student_info: Dict, behavior_defs: List[Dict]) -> Dict:
        """
        生成 BIP 计划：包含分类定义、FERB、Prompt 消退路径与 Staff/Paraprofessional Training
        """
        bip_modules = []
        
        for beh in behavior_defs:
            bip_modules.append({
                "target_behavior": beh["name"],
                "operational_definition": beh["definition"],  # BIP 开头保留精简定义
                "function": beh["function"],
                "ferb": beh["replacement_behavior"],
                "prompting_strategy": {
                    "method": "Systematic Prompt Fading Plan",
                    "hierarchy": [
                        "1. Full Physical Prompt (完全身体引导)",
                        "2. Partial Physical Prompt (部分身体引导)",
                        "3. Visual / Gestural Prompt (视觉/手势提示)",
                        "4. Independent (独立完成替代行为)"
                    ],
                    "fading_criterion": "在连续 3 天的观察中达到 80% 正确独立反应后降级提示。"
                }
            })
            
        # BST 员工培训与监控模块
        staff_training_protocol = {
            "target_audience": "Teachers, Paraprofessionals, and Educational Assistants",
            "model": "Behavioral Skills Training (BST)",
            "steps": [
                "1. Instruction: 详细讲解 BIP 操作流程及 FERB 的强化时机。",
                "2. Modeling: 督导 (BCBA/BSP) 在真实课堂环境中现场示范干预流程。",
                "3. Rehearsal: 助教/教师在角色扮演或真实环境中演练干预。",
                "4. Feedback: 提供即时的肯定与纠正性反馈。"
            ],
            "fidelity_monitoring": {
                "check_frequency": "每周一次（初期），稳定后每两周一次",
                "tool": "Treatment Fidelity Checklist (涵盖 Prompt 提示与延迟强化执行准确度)",
                "self_monitoring": "助教每日课程结束时填写简易 Self-Monitoring Sheet。"
            }
        }
        
        return {
            "identifying_data": student_info,
            "behavior_interventions": bip_modules,
            "staff_training_and_monitoring": staff_training_protocol
        }

# ==================== 示范运行 ====================
if __name__ == "__main__":
    engine = FBABIPEngine(language="zh")
    
    # 1. 模拟儿童基本信息
    student_data = {
        "name": "Child A",
        "age": "3.5",
        "setting": "Early Intervention Classroom",
        "antecedent_trigger": "高认知负荷任务或感官过载环境"
    }
    
    # 2. 测试 QABF 与 ABC 数据不一致的情况
    qabf_input = {"Attention": 12, "Escape": 15, "Sensory": 8}
    abc_input = {"Attention": 2, "Escape": 18, "Sensory": 3}  # ABC 强烈指向 Escape
    
    # 计算冲突处理
    function_res = engine.resolve_function_discrepancy(qabf_input, abc_input)
    
    # 3. 生成 FBA 第5点（Synthesis）
    fba_sec_5 = engine.generate_fba_section_5(
        student_info=student_data,
        target_behavior="Face Slapping (Target Behavior 2)",
        function_analysis=function_res
    )
    
    # 4. 生成 BIP 完整方案
    behaviors = [
        {
            "name": "Face Slapping",
            "definition": "使用单手或双手掌心拍打自己面部，产生清晰可听见的响声或皮肤发红。",
            "function": function_res["primary_function"],
            "replacement_behavior": "使用 Sensory Break Button 或出示 'I need a break' 视觉卡片"
        }
    ]
    bip_res = engine.generate_bip_plan(student_data, behaviors)
    
    # 打印测试输出
    print("=== FBA Section 5: Individualized Synthesis ===")
    print(json.dumps(fba_sec_5, ensure_ascii=False, indent=2))
    
    print("\n=== BIP Section: Staff Training Protocol ===")
    print(json.dumps(bip_res["staff_training_and_monitoring"], ensure_ascii=False, indent=2))
