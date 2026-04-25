"""
CLI entry point.

Usage:
    python main.py --input project.xer --output dashboard_data.xlsx
    python main.py --sample --output dashboard_data.xlsx
"""

import sys
from pathlib import Path

import click

from xer_parser       import parse_xer
from sample_generator import generate_sample_xer
from data_model       import process
from excel_exporter   import export_excel


@click.command()
@click.option("--input",  "-i", "input_path",  type=click.Path(), default=None,
              help="Path to Primavera P6 .XER file")
@click.option("--output", "-o", "output_path", type=click.Path(), default="dashboard_data.xlsx",
              show_default=True, help="Output Excel file path")
@click.option("--sample", is_flag=True, default=False,
              help="Generate and parse a sample XER instead of reading a file")
def main(input_path: str | None, output_path: str, sample: bool) -> None:
    """XER → Power BI Excel exporter for Primavera P6 projects."""

    if sample:
        xer_path = Path(output_path).with_suffix(".xer")
        generate_sample_xer(xer_path)
        input_path = str(xer_path)
        click.echo(f"[main] Sample XER generated: {xer_path}")

    if not input_path:
        click.echo("ERROR: provide --input <file.xer> or use --sample", err=True)
        sys.exit(1)

    click.echo(f"[main] Parsing: {input_path}")
    tables = parse_xer(input_path)
    click.echo(f"[main] Tables found: {list(tables.keys())}")

    click.echo("[main] Processing data model...")
    datasets = process(tables)

    for name, df in datasets.items():
        click.echo(f"  {name:15s}: {len(df):>5} rows × {len(df.columns):>3} cols")

    click.echo(f"[main] Exporting Excel → {output_path}")
    export_excel(datasets, output_path)
    click.echo("[main] Done.")


if __name__ == "__main__":
    main()
