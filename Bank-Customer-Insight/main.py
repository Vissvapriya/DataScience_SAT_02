"""Run all four stages of the TrustWave Bank synthetic-data project."""

from __future__ import annotations

import json

from notebooks.round1_eda import run_round1
from notebooks.round2_machine_learning import run_round2
from notebooks.round3_visualization import run_round3


def main() -> None:
    print("Round 1/3: generating synthetic data, cleaning, and creating EDA...")
    findings = run_round1()
    print(
        f"  {findings['source_record_counts']['customers']:,} synthetic customers; "
        f"top channel: {findings['top_channel']}."
    )
    print("Round 2/3: engineering features and evaluating models...")
    model_results = run_round2()
    selected = model_results["models"][model_results["selected_model"]]
    print(
        f"  Selected {model_results['selected_model']}: "
        f"accuracy={selected['accuracy']:.4f}, macro F1={selected['macro_f1']:.4f}."
    )
    print("Round 3/3: building Excel, reports, and presentation...")
    deliverables = run_round3()
    print(json.dumps(deliverables, indent=2))
    print("Workflow complete. All data and results are synthetic.")


if __name__ == "__main__":
    main()
