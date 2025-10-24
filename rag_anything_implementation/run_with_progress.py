#!/usr/bin/env python3
import argparse
import os
import re
import subprocess
import sys
import time
from threading import Thread
from queue import Queue, Empty

# No page counting - just show elapsed time


def enqueue_output(stream, queue: Queue, tag: str):
    for line in iter(stream.readline, b''):
        queue.put((tag, line))
    stream.close()


def human_time(seconds: float) -> str:
    seconds = max(0, int(seconds))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h:d}h {m:02d}m {s:02d}s"
    if m:
        return f"{m:d}m {s:02d}s"
    return f"{s:d}s"


def render_timer(prefix: str, elapsed: float, last_activity: str = "") -> str:
    activity_info = f" | {last_activity}" if last_activity else ""
    return f"\r{prefix} | elapsed {human_time(elapsed)}{activity_info}"


def get_last_activity(output_dir: str) -> str:
    """Get a simple status based on output directory contents"""
    if not os.path.exists(output_dir):
        return "Starting..."
    
    files = os.listdir(output_dir)
    
    # Look for common intermediate files
    if any(f.endswith('.json') for f in files):
        return "Processing content..."
    if any(f.endswith('.pdf') for f in files):
        return "Generating layout..."
    if any('auto' in f for f in files):
        return "Finalizing output..."
    
    return "Initializing..."


def main():
    parser = argparse.ArgumentParser(description="Run end-2-end RAGAnything with elapsed timer")
    parser.add_argument("--pdf", default=os.path.join(os.path.dirname(__file__), "input", "fullsheets-vba.pdf"), help="Path to the input PDF")
    parser.add_argument("--script", default=os.path.join(os.path.dirname(__file__), "end-2-end-rag-anything.py"), help="Path to the end-to-end script to run")
    parser.add_argument("--quiet-child", action="store_true", help="Do not mirror child stdout/stderr (only show timer)")
    parser.add_argument("--output-dir", default="output", help="Output directory to monitor for status")
    args, unknown = parser.parse_known_args()

    pdf_path = os.path.abspath(args.pdf)
    script_path = os.path.abspath(args.script)

    if not os.path.isfile(pdf_path):
        print(f"Error: PDF not found at {pdf_path}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(script_path):
        print(f"Error: script not found at {script_path}", file=sys.stderr)
        sys.exit(1)

    # No page counting needed

    # Export environment to ensure the child uses the same PDF (override inside the script if needed)
    env = os.environ.copy()
    env["RAG_INPUT_PDF"] = pdf_path

    cmd = [sys.executable, script_path]
    if unknown:
        cmd.extend(unknown)

    start = time.time()
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)

    q: Queue = Queue()
    t_out = Thread(target=enqueue_output, args=(proc.stdout, q, 'O'))
    t_err = Thread(target=enqueue_output, args=(proc.stderr, q, 'E'))
    t_out.daemon = True
    t_err.daemon = True
    t_out.start()
    t_err.start()

    last_render = ''
    last_activity = ""
    last_log_time = start
    try:
        while True:
            try:
                tag, raw = q.get(timeout=0.05)
                try:
                    line = raw.decode(errors='replace')
                except Exception:
                    line = str(raw)
                
                # Update activity based on log content
                if "Detected PDF file, using parser for PDF" in line:
                    last_activity = "PDF detected, starting parser..."
                elif "Processing page" in line or "Page" in line:
                    # Just show we're processing pages
                    last_activity = "Processing pages..."
                elif "parsing" in line.lower() and ("complete" in line.lower() or "finished" in line.lower()):
                    last_activity = "Parsing complete"
                elif "query" in line.lower() and "result" in line.lower():
                    last_activity = "Running queries..."
                
                # Update last log time when we get new output
                last_log_time = time.time()
                
                if not args.quiet_child:
                    if tag == 'O':
                        sys.stdout.write(line)
                        sys.stdout.flush()
                    else:
                        sys.stderr.write(line)
                        sys.stderr.flush()
            except Empty:
                pass

            if proc.poll() is not None:
                break

            elapsed = time.time() - start
            
            # Get current activity status
            current_activity = get_last_activity(args.output_dir)
            if current_activity != "Starting...":
                last_activity = current_activity
            
            # Show timer with last activity
            render = render_timer("Processing PDF", elapsed, last_activity)
            if render != last_render:
                sys.stdout.write(render)
                sys.stdout.flush()
                last_render = render
        # Child finished
        elapsed = time.time() - start
        render = render_timer("Processing PDF", elapsed, "Complete")
        sys.stdout.write(render + "\n")
        sys.stdout.flush()
    finally:
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()

    sys.exit(proc.returncode)


if __name__ == "__main__":
    main()
