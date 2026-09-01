from datetime import datetime

import streamlit as st

from Cal_TPH_RespTime_jtl import calculate_response_jtl


def show_jtl_preview(
    jtl_files,
    from_date,
    from_time,
    to_date,
    to_time,
    keywords=None,
    template_file=None,
    data=None,
    template=None,
):
    """Display a preview of the filtered JTL summary with total iterations and failures."""
    try:
        from_dt = datetime.combine(from_date, from_time)
        to_dt = datetime.combine(to_date, to_time)
        template_input = template if template is not None else template_file

        total_summary = calculate_response_jtl(
            jtl_files=jtl_files,
            from_date=from_dt,
            to_date=to_dt,
            pass_only=False,
            keywords=keywords,
            template=template_input,
            data=data,
        )
        pass_summary = calculate_response_jtl(
            jtl_files=jtl_files,
            from_date=from_dt,
            to_date=to_dt,
            pass_only=True,
            keywords=keywords,
            template=template_input,
            data=data,
        )

        preview_summary = pass_summary.copy()
        preview_summary = preview_summary.rename(
            columns={
                "label": "Transaction Name",
                "Samples": "Pass",
                "PassPercent": "Pass %",
            }
        )

        iteration_counts = total_summary[["label", "Samples"]].rename(
            columns={"label": "Transaction Name", "Samples": "Iterations"}
        )
        preview_summary = preview_summary.merge(iteration_counts, on="Transaction Name", how="left")

        preview_summary["Pass"] = preview_summary["Pass"].fillna(0).astype(int)
        preview_summary["Iterations"] = preview_summary["Iterations"].fillna(0).astype(int)
        preview_summary["Fail"] = (preview_summary["Iterations"] - preview_summary["Pass"]).clip(lower=0).astype(int)

        if "Pass %" in preview_summary.columns:
            preview_summary["Pass %"] = (
                preview_summary["Pass"] / preview_summary["Iterations"].replace(0, float("nan")) * 100
            ).fillna(0).round(3)

        preview_summary = preview_summary[
            [
                "Transaction Name",
                "Iterations",
                "Pass",
                "Fail",
                "Pass %",
                *[col for col in preview_summary.columns if col not in {"Transaction Name", "Iterations", "Pass", "Fail", "Pass %", "Throughput"}],
            ]
        ]
        st.dataframe(preview_summary, width='stretch')
    except Exception as error:
        st.error(f"Unable to preview JTL data: {error}")
