import pandas as pd
import pytest


@pytest.fixture
def plan_df_factory(tmp_path):
    def make(output_names):
        rows = []
        for index, output_name in enumerate(output_names):
            src = tmp_path / f"src_{index}.txt"
            src.write_text("x")
            rows.append({
                "原文件名": src.name,
                "输出文件名": output_name,
                "原扩展名": ".txt",
                "新文件名主体": output_name,
                "清洗后的新文件名主体": output_name,
                "是否保留扩展名": True,
                "source_path": str(src),
            })
        return pd.DataFrame(rows)
    return make
