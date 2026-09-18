import argparse
import io
import sys
import time
from pathlib import Path
from dotenv import load_dotenv

# Ensure safe UTF-8 output on Windows console
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Load environment variables
load_dotenv()

from certificate_agent.logger import logger
from certificate_agent.tools import (
    validate_inputs,
    generate_all_certificates,
    detect_excel_changes,
    process_excel_changes,
)
from certificate_agent.file_watcher import ExcelMonitor


def cmd_validate(args):
    """Validate Excel dataset and certificate template."""
    print("\n--- Validating Inputs ---")
    res = validate_inputs(excel_path=args.excel, template_path=args.template)
    if res["valid"]:
        print(f"✓ Inputs are valid!")
        print(f"  Excel File: {res['excel_path']}")
        print(f"  Template File: {res['template_path']}")
        print(f"  Student Records: {res['records_count']}")
        print(f"  Template Placeholders: {', '.join(res['placeholders'])}")
        if res["warnings"]:
            print("\nWarnings:")
            for w in res["warnings"]:
                print(f"  ! {w}")
    else:
        print("✗ Validation failed:")
        for err in res["errors"]:
            print(f"  ✗ {err}")
        sys.exit(1)


def cmd_generate(args):
    """Generate certificates for all students."""
    try:
        res = generate_all_certificates(
            excel_path=args.excel,
            template_path=args.template,
            output_dir=args.output
        )
        print("\n" + res["summary"])
    except Exception as e:
        logger.error(f"Generation failed: {e}")
        print(f"\nGeneration Error: {e}", file=sys.stderr)
        sys.exit(1)


def cmd_sync(args):
    """Synchronize certificates with Excel changes."""
    try:
        res = process_excel_changes(
            excel_path=args.excel,
            delete_removed=args.delete_removed
        )
        print("\n" + res["summary"])
    except Exception as e:
        logger.error(f"Sync failed: {e}")
        print(f"\nSync Error: {e}", file=sys.stderr)
        sys.exit(1)


def cmd_watch(args):
    """Continuously monitor Excel file for changes."""
    print("\n--- Starting Excel Change Monitor ---")
    print("Press Ctrl+C to stop monitoring.\n")

    def on_excel_changed():
        print("\n[Change Detected] Updating certificates...")
        try:
            res = process_excel_changes(excel_path=args.excel, delete_removed=args.delete_removed)
            print(res["summary"])
        except Exception as e:
            print(f"Error updating certificates: {e}", file=sys.stderr)

    monitor = ExcelMonitor(
        excel_path=Path(args.excel) if args.excel else None,
        debounce_seconds=args.debounce,
        on_change_callback=on_excel_changed
    )

    monitor.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping monitor...")
        monitor.stop()
        print("Monitor stopped.")


def cmd_agent(args):
    """Run interactive ADK agent or verify agent readiness."""
    from certificate_agent.agent import root_agent
    print("\n==============================================")
    print("       ADK Certificate Generation Agent       ")
    print("==============================================")
    print(f"Agent Name: {root_agent.name}")
    print(f"Model: {root_agent.model}")
    print(f"Tools Registered: {len(root_agent.tools)}")
    for t in root_agent.tools:
        name = getattr(t, "__name__", str(t))
        print(f"  - {name}")
    print("\nAgent initialized successfully and ready for ADK orchestration.")


def main():
    parser = argparse.ArgumentParser(
        description="ADK Certificate Generator with Excel Change Detection"
    )
    parser.add_argument("--excel", help="Path to student Excel file")
    parser.add_argument("--template", help="Path to certificate DOCX template")
    parser.add_argument("--output", help="Path to output directory for PDFs")

    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # validate
    sub_val = subparsers.add_parser("validate", help="Validate Excel and template")
    sub_val.set_defaults(func=cmd_validate)

    # generate
    sub_gen = subparsers.add_parser("generate", help="Generate all student certificates")
    sub_gen.set_defaults(func=cmd_generate)

    # sync
    sub_sync = subparsers.add_parser("sync", help="Sync Excel changes (process only new/modified)")
    sub_sync.add_argument("--delete-removed", action="store_true", default=None,
                           help="Delete PDFs of students removed from Excel")
    sub_sync.set_defaults(func=cmd_sync)

    # watch
    sub_watch = subparsers.add_parser("watch", help="Monitor Excel file and sync continuously")
    sub_watch.add_argument("--debounce", type=float, default=2.0, help="Debounce seconds (default 2.0)")
    sub_watch.add_argument("--delete-removed", action="store_true", default=None,
                            help="Delete PDFs of students removed from Excel")
    sub_watch.set_defaults(func=cmd_watch)

    # agent
    sub_agent = subparsers.add_parser("agent", help="Display ADK agent info / run agent")
    sub_agent.set_defaults(func=cmd_agent)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    args.func(args)


if __name__ == "__main__":
    main()
