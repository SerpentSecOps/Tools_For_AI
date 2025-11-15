#!/usr/bin/env python3
#Unified LLM Content Preparation Tool
#Combines project code bundling and document knowledge file creation
#for optimal LLM processing


import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
import os
import sys
import fnmatch
import hashlib
import datetime
import concurrent.futures
import queue
import threading
import multiprocessing
import shutil
import unicodedata
import re
from pathlib import Path
from typing import List, Iterable, Callable, Dict, Any, Tuple, Optional

# Third-party imports (for document processing)
try:
    import fitz  # PyMuPDF
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    from bs4.element import NavigableString
    import mobi
    DOCUMENT_SUPPORT = True
except ImportError:
    DOCUMENT_SUPPORT = False

# =======================================================================================
# SHARED UTILITIES
# =======================================================================================

def sanitize_text(s: str) -> str:
    """Normalize unicode, standardize newlines, strip harmful control chars (keep \n and \t)."""
    s = unicodedata.normalize('NFKC', s)
    s = s.replace('\r\n', '\n').replace('\r', '\n')
    # Remove control characters except tabs and newlines
    control_filter = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")
    s = control_filter.sub('', s)
    return s

def stable_hash(content: str) -> str:
    """Generate a stable SHA256 hash for content."""
    return f"sha256-{hashlib.sha256(content.encode('utf-8')).hexdigest()}"

# =======================================================================================
# PROJECT CODE BUNDLER (from code_manifest.py)
# =======================================================================================

DEFAULT_IGNORE = [
    # VCS
    ".git/", ".gitignore", ".gitattributes",
    # Common Lockfiles
    "poetry.lock", "pnpm-lock.yaml", "package-lock.json", "yarn.lock",
    # Python
    "__pycache__/", "*.pyc", "*.pyo", "*.pyd", "*.egg", "*.egg-info/", "pip-wheel-metadata/",
    # Virtual Environments
    "venv/", ".venv/", "env/", ".tox/",
    # Dotnet
    "bin/", "obj/", "*.csproj.user", "*.sln.dotsettings",
    # Node
    "node_modules/", ".pnpm-store/",
    # Env
    ".env", ".env.*",
    # IDE
    "nbproject/", "*.sublime-workspace", ".vscode/", ".idea/",
    # PHP
    "vendor/",
    # Build artifacts
    "build/", "dist/", "target/", "out/",
    # Logs, DBs, caches
    "*.log", "*.db", "*.sqlite", "*.sqlite3", "*.db-journal",
    # OS-specific
    ".DS_Store", "Thumbs.db",
]

TEXT_EXT_HINT = {
    # Code-ish
    ".py": "Python", ".pyi": "Python Stub", ".ipynb": "Jupyter Notebook",
    ".js": "JavaScript", ".jsx": "JavaScript (React)", ".mjs": "JavaScript Module",
    ".ts": "TypeScript", ".tsx": "TypeScript (React)",
    ".c": "C", ".h": "C Header", ".cpp": "C++", ".hpp": "C++ Header", ".cc": "C++",
    ".rs": "Rust", ".go": "Go", ".java": "Java", ".kt": "Kotlin", ".kts": "Kotlin Script", ".scala": "Scala",
    ".rb": "Ruby", ".php": "PHP", ".swift": "Swift", ".cs": "C#",
    ".m": "Objective-C", ".mm": "Objective-C++",
    ".sh": "Shell", ".bash": "Shell", ".zsh": "Shell", ".fish": "Shell", ".ps1": "PowerShell",
    # Web/markup/config
    ".html": "HTML", ".htm": "HTML", ".css": "CSS", ".scss": "SCSS",
    ".json": "JSON", ".jsonc": "JSON with Comments",
    ".yml": "YAML", ".yaml": "YAML", ".toml": "TOML", ".ini": "INI",
    ".md": "Markdown", ".rst": "reStructuredText", ".sql": "SQL",
    ".xml": "XML", ".xsl": "XSLT", ".xslt": "XSLT", ".svg": "SVG",
    ".dockerfile": "Dockerfile", "Dockerfile": "Dockerfile", ".env": "Env",
    # Data-ish
    ".csv": "CSV", ".tsv": "TSV", ".txt": "Text", ".log": "Log",
    # Docs & Other
    ".tex": "LaTeX", ".cls": "LaTeX", ".sty": "LaTeX",
    # Templates
    ".jinja": "Jinja", ".jinja2": "Jinja", ".tmpl": "Template",
}

BINARY_EXT_LIKELY = {
    ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".ico",
    ".pdf", ".zip", ".gz", ".tar", ".tgz", ".bz2", ".xz", ".7z",
    ".so", ".dll", ".dylib", ".exe", ".bin", ".class", ".o", ".a",
    ".ttf", ".otf", ".woff", ".woff2",
    ".mp3", ".wav", ".flac", ".ogg", ".mp4", ".mov", ".mkv", ".avi",
}

def read_gitignore_patterns(project_root: Path) -> List[str]:
    patterns = []
    gi = project_root / ".gitignore"
    if gi.exists():
        try:
            with gi.open("r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        patterns.append(line)
        except Exception:
            pass
    return patterns

def norm_for_match(p: Path, root: Path) -> str:
    """Forward-slash, root-relative path; dirs end with '/'."""
    try:
        rel = p.relative_to(root)
    except Exception:
        rel = p
    s = str(rel).replace("\\", "/")
    if p.is_dir() and not s.endswith("/"):
        s += "/"
    return s

def should_ignore(path: Path, root: Path, patterns: List[str]) -> bool:
    """Match against both the name and normalized root-relative path."""
    name = path.name
    rel = norm_for_match(path, root)
    for pat in patterns:
        pat = pat.replace("\\", "/")
        if pat.endswith("/"):
            if fnmatch.fnmatch(name + "/", pat) or fnmatch.fnmatch(rel, pat) or rel.startswith(pat):
                return True
        else:
            if fnmatch.fnmatch(name, pat) or fnmatch.fnmatch(rel, pat):
                return True
    return False

def detect_language(path: Path) -> str:
    return TEXT_EXT_HINT.get(path.suffix.lower(), "Plain Text")

def is_probably_binary(path: Path) -> bool:
    ext = path.suffix.lower()
    if ext in BINARY_EXT_LIKELY:
        return True
    try:
        with path.open("rb") as f:
            chunk = f.read(4096)
        if b"\x00" in chunk:
            return True
    except Exception:
        return True
    return False

def iter_files(root: Path, patterns: List[str], follow_symlinks: bool) -> Iterable[Path]:
    for cur_root, dirs, files in os.walk(root, topdown=True, followlinks=follow_symlinks):
        cur_root_p = Path(cur_root)
        dirs[:] = sorted([d for d in dirs if not should_ignore(cur_root_p / d, root, patterns)])
        for fname in sorted(files):
            p = cur_root_p / fname
            if should_ignore(p, root, patterns):
                continue
            yield p

def make_tree_map(root: Path, visible_files: List[Path]) -> str:
    """Render a compact project tree of visible files/dirs."""
    visible_set = set(visible_files)
    dir_set = {root}
    for f in visible_files:
        for parent in f.parents:
            if parent == parent.parent:
                break
            if root in parent.parents or parent == root:
                dir_set.add(parent)

    lines = []
    lines.append(f"{root.name or str(root)}/")
    children = {d: {"dirs": [], "files": []} for d in dir_set}

    for f in visible_files:
        parent = f.parent
        if parent in children:
            children[parent]["files"].append(f)
    for d in list(dir_set):
        parent = d.parent
        if parent in children and d != root:
            children[parent]["dirs"].append(d)

    def render(d: Path, prefix: str = ""):
        subdirs = sorted(children.get(d, {}).get("dirs", []), key=lambda x: x.name)
        files = sorted(children.get(d, {}).get("files", []), key=lambda x: x.name)
        for i, sd in enumerate(subdirs):
            is_last = (i == len(subdirs) - 1) and not files
            branch = "└── " if is_last else "├── "
            lines.append(prefix + branch + sd.name + "/")
            new_prefix = prefix + ("    " if is_last else "│   ")
            render(sd, new_prefix)
        for j, f in enumerate(files):
            is_last = j == len(files) - 1
            branch = "└── " if is_last else "├── "
            lines.append(prefix + branch + f.name)

    render(root)
    return "\n".join(lines)

def render_llm_usage_guide(guide_mode: str) -> str:
    if guide_mode == "none":
        return ""
    base = []
    base.append("## LLM USAGE GUIDE")
    base.append("")
    base.append("### QUICKSTART")
    base.append("- **Search by file ID** (e.g., `F0007`) for an exact match. IDs are stable across runs if paths don't change.")
    base.append("- Use the **FILE INDEX** table below to find IDs, paths, languages, byte sizes, and line counts.")
    base.append("- Each file section is delimited by clear markers:")
    base.append("  - `===== FILE FXXXX =====` (header and metadata)")
    base.append("  - `----- BEGIN CONTENT FXXXX -----`")
    base.append("  - `----- END CONTENT FXXXX -----`")
    base.append("- If a file shows `NOTE: truncated ...`, ask the user for the original file if needed.")
    base.append("- Binary files are **skipped** with a clear note to avoid parsing noise.")
    base.append("")
    base.append("### SEARCH TIPS")
    base.append("- Prefer `FXXXX` IDs over ambiguous names like `control_panel`.")
    base.append("- To anchor to a path, search for `PATH: some/dir/file.py` within file headers.")
    base.append("- To jump through files quickly, grep for `===== FILE F` markers.")
    base.append("")
    if guide_mode == "verbose":
        base.append("### READING STRATEGY (FOR SMALL CONTEXT MODELS)")
        base.append("1) Read **PROJECT MAP** to understand structure.")
        base.append("2) Scan **FILE INDEX** to pick likely targets by path/language/size.")
        base.append("3) Open the **smallest relevant files first** to save context.")
        base.append("4) Use IDs consistently in your notes/responses (e.g., 'Changes in `F0012`').")
        base.append("5) If instructions mention 'don't modify detection logic', keep logic untouched and add UI-only changes.")
        base.append("")
        base.append("### WHEN YOU NEED MORE CONTEXT")
        base.append("- If a file is truncated or missing, mention the `ID` and ask the user for the original file.")
        base.append("- If multiple files reference the same concept, list the relevant IDs before summarizing.")
        base.append("")
    base.append("### PROMPT TEMPLATES")
    base.append("- **Locate a file quickly:**")
    base.append('  - "Find `F0012` and summarize its purpose in 3 bullets."')
    base.append("- **Apply a patch safely:**")
    base.append('  - "Open `F0012` (`PATH: gui/control_panel.py`). Add a checkbox + spinbox UI (no changes to detection logic). Preserve everything else. Provide a unified diff."')
    base.append("- **Cross-file question:**")
    base.append('  - "Which files import `ModelLoader`? Return IDs and a one-line description for each."')
    base.append("")
    return "\n".join(base) + "\n\n"

# =======================================================================================
# DOCUMENT KNOWLEDGE FILE CREATOR (from gem_convert.py)
# =======================================================================================

if DOCUMENT_SUPPORT:
    KNOWLEDGE_FILE_HEADER = """
[SYSTEM INSTRUCTION]
This is a structured knowledge file. Interpret it according to these rules:
1.  **File Structure:** Begins with a Table of Contents (TOC).
2.  **Document ID (DocID):** Each document has a short, unique `DocID` for citation.
3.  **Content Hash:** A full SHA256 hash is provided for data integrity.
4.  **Markers:** Content is encapsulated by `[START/END OF DOCUMENT]` markers.
5.  **Usage:** Use the content to answer queries, citing the `DocID` and Title.
[/SYSTEM INSTRUCTION]
---
"""

    REFERENCE_SHEET_HEADER = """
[SYSTEM INSTRUCTION]
This is an AI Reference Sheet, a global manifest for multiple knowledge files.
1.  **Purpose:** This file is an index linking a document's `DocID` and `Title` to its `SourceFile`. It does not contain document text.
2.  **Structure:** `[DocID: ... | Title: ...] | [SourceFile: ...]`
3.  **Usage:** Use this manifest to identify which `SourceFile` contains the relevant document for a query before retrieving the content.
4.  **Canonical Identifier:** The `DocID` is the unique identifier.
[/SYSTEM INSTRUCTION]
---
"""

    # Control characters to filter (keep tabs/newlines)
    _CONTROL_FILTER = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")

    def _normalize_prose(s: str) -> str:
        """Collapse excessive spaces while preserving line structure."""
        s = re.sub(r"[ \t]+(?=\n)", '', s)  # Trim trailing spaces at end of lines
        s = re.sub(r"[ ]{2,}", ' ', s)      # Collapse multiple spaces inside lines
        s = re.sub(r"\n{3,}", "\n\n", s)   # Reduce 3+ blank lines to 2
        return s.strip()

    CODE_START = "@@@CODEBLOCK_START@@@"
    CODE_END = "@@@CODEBLOCK_END@@@"
    _CODE_OR_MARKERS = re.compile(r"(?s)(```.*?```|" + re.escape(CODE_START) + r".*?" + re.escape(CODE_END) + r")")

    def normalize_text_code_safe(s: str) -> str:
        """Sanitize, then normalize prose but keep code blocks verbatim."""
        s = sanitize_text(s)
        out = []
        i = 0
        for m in _CODE_OR_MARKERS.finditer(s):
            pre = s[i:m.start()]
            code = m.group(0)
            if pre:
                out.append(_normalize_prose(pre))
            out.append(code)  # keep code untouched
            i = m.end()
        tail = s[i:]
        if tail:
            out.append(_normalize_prose(tail))
        return ''.join(out)

    def _html_to_text_preserving_code(html: str) -> str:
        soup = BeautifulSoup(html, 'html.parser')
        for tag in soup.find_all(['pre', 'code', 'samp', 'kbd']):
            tag.insert_before(NavigableString(CODE_START))
            tag.insert_after(NavigableString(CODE_END))
        txt = soup.get_text(separator='\n')
        return normalize_text_code_safe(txt)

    def _normalize_title_from_path(file_path: str) -> str:
        raw_name = os.path.splitext(os.path.basename(file_path))[0]
        safe = re.sub(r"[^\w\s\-\.,'()&]+", ' ', raw_name).strip()
        return normalize_text_code_safe(safe.title())

    def _stable_sample(text: str, k: int = 2000) -> str:
        n = len(text)
        if n <= k:
            return text
        mid = n // 2
        parts = [text[:k], text[max(0, mid - k // 2): mid + k // 2], text[-k:]]
        return ''.join(parts)

    class IDManager:
        ALPHABET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        @staticmethod
        def _int_to_base_n(n: int, base: int) -> str:
            if n == 0:
                return IDManager.ALPHABET[0]
            s = []
            while n:
                s.append(IDManager.ALPHABET[n % base])
                n //= base
            return "".join(reversed(s))

        @staticmethod
        def generate_short_id(content_hash: str, prefix: str = "DOC", length: int = 6) -> str:
            try:
                hash_hex = content_hash.split('-', 1)[1][:12]
            except Exception:
                hash_hex = hashlib.sha256(content_hash.encode('utf-8')).hexdigest()[:12]
            hash_int = int(hash_hex, 16)
            base36_id = IDManager._int_to_base_n(hash_int, len(IDManager.ALPHABET))
            return f"{prefix}{base36_id.zfill(length)}"

    class ExtractionError(Exception):
        pass

    # Document extractors
    def extract_text_from_txt(file_path: str) -> str:
        try:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    raw = f.read()
            except (UnicodeDecodeError, IOError):
                with open(file_path, 'r', encoding='latin-1', errors='ignore') as f:
                    raw = f.read()
            text = normalize_text_code_safe(raw)
            if not text.strip():
                raise ExtractionError("Empty text after normalization.")
            return text
        except Exception as e:
            raise ExtractionError(f"Reason: {e}") from e

    def extract_text_from_pdf(file_path: str) -> str:
        try:
            with fitz.open(file_path) as doc:
                text = "".join(page.get_text("text") for page in doc)
            text = normalize_text_code_safe(text)
            if not text:
                raise ExtractionError("No selectable text found (likely scanned PDF).")
            return text
        except Exception as e:
            raise ExtractionError(f"Reason: {e}") from e

    def extract_text_from_epub(file_path: str) -> str:
        try:
            book = epub.read_epub(file_path)
            parts = []
            for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
                name = (item.get_name() or '').lower()
                if name.endswith('.css'):
                    continue
                t = _html_to_text_preserving_code(item.get_body_content())
                if t.strip():
                    parts.append(t)
            if not parts:
                raise ExtractionError("No text documents found in EPUB.")
            return "\n\n".join(parts)
        except Exception as e:
            raise ExtractionError(f"Reason: {e}") from e

    def extract_text_from_mobi(file_path: str) -> str:
        temp_dir = None
        try:
            temp_dir, _ = mobi.extract(file_path)
            parts = []
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    if file.lower().endswith(('.html', '.htm', '.txt')):
                        with open(os.path.join(root, file), 'r', encoding='utf-8', errors='ignore') as f:
                            parts.append(_html_to_text_preserving_code(f.read()))
            if not parts:
                raise ExtractionError("No text content found after MOBI unpack.")
            return "\n\n".join(parts)
        except Exception as e:
            raise ExtractionError(f"Reason: {e}") from e
        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

    def process_single_file(file_path: str, id_prefix: str) -> Tuple[str, Optional[Dict[str, Any]]]:
        filename = os.path.basename(file_path)
        title = _normalize_title_from_path(file_path)
        ext = os.path.splitext(filename)[1].lower()

        extractor_map: Dict[str, Callable[[str], str]] = {
            '.txt': extract_text_from_txt,
            '.md': extract_text_from_txt,
            '.markdown': extract_text_from_txt,
            '.rst': extract_text_from_txt,
            '.csv': extract_text_from_txt,
            '.tsv': extract_text_from_txt,
            '.log': extract_text_from_txt,
            '.json': extract_text_from_txt,
            '.xml': extract_text_from_txt,
            '.yaml': extract_text_from_txt,
            '.yml': extract_text_from_txt,
            '.toml': extract_text_from_txt,
            '.ini': extract_text_from_txt,
            '.cfg': extract_text_from_txt,
            '.conf': extract_text_from_txt,
            '.sql': extract_text_from_txt,
            '.tex': extract_text_from_txt,
            '.rtf': extract_text_from_txt,
            '.pdf': extract_text_from_pdf,
            '.epub': extract_text_from_epub,
            '.mobi': extract_text_from_mobi,
        }

        if ext not in extractor_map:
            return file_path, {"error": f"Unsupported file type"}

        try:
            text = extractor_map[ext](file_path)
            if not text.strip():
                return file_path, {"error": "Extracted text is empty."}

            hash_content = normalize_text_code_safe(title) + _stable_sample(text, 2000)
            full_hash = stable_hash(hash_content)
            short_id = IDManager.generate_short_id(full_hash, prefix=id_prefix, length=6)

            return file_path, {
                "short_id": short_id,
                "full_hash": full_hash,
                "title": title,
                "text": text,
            }
        except ExtractionError as e:
            return file_path, {"error": f"{type(e).__name__}: {e}"}

# =======================================================================================
# UNIFIED GUI APPLICATION
# =======================================================================================

HELP_TEXT = """
**Unified LLM Content Preparation Tool**

This tool combines two powerful features for preparing content for LLMs:

**1. PROJECT CODE BUNDLER**
- Bundles entire code projects into a single, well-structured file
- Includes project tree, file index with stable IDs, and LLM usage guide
- Perfect for code analysis, debugging, and development tasks
- Respects .gitignore and includes comprehensive ignore patterns

**2. DOCUMENT KNOWLEDGE FILE CREATOR** (requires PyMuPDF, ebooklib, beautifulsoup4, mobi)
- Converts books and documents (TXT, PDF, EPUB, MOBI) into structured knowledge files
- Creates reference sheets for managing multiple knowledge bases
- Optimized for AI training and knowledge retrieval tasks
- Handles batch processing with progress tracking

**Usage:**
1. Choose your preparation method using the tabs
2. Configure settings and select input files/directories  
3. Process your content and get LLM-ready output files

**Benefits:**
- Stable, searchable file IDs for consistent referencing
- Structured output optimized for LLM token efficiency
- Comprehensive metadata and indexing
- Handles large projects and document collections
"""

ABOUT_TEXT = """
**Unified LLM Content Preparation Tool v1.0**

A comprehensive solution for preparing content for Large Language Models.

**Features:**
• Project code bundling with intelligent file detection
• Document knowledge file creation and management
• Batch processing with progress tracking
• LLM-optimized output formatting
• Duplicate detection and queue management
• Comprehensive file type support

**Author:** Combined from code_manifest.py and gem_convert.py
**Purpose:** Streamline LLM content preparation workflows
"""

class UnifiedLLMPrepTool:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Unified LLM Content Preparation Tool v1.0")
        self.root.geometry("1000x900")
        
        # Shared state
        self.log_queue: "queue.Queue[Tuple[str, str]]" = queue.Queue()
        self.processing_thread: Optional[threading.Thread] = None
        
        # Document processor state
        self.file_queue: List[str] = []
        self.docid_prefix: str = "DOC"
        
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Check for document processing support
        if not DOCUMENT_SUPPORT:
            self.log_message("WARNING: Document processing libraries not found. Install PyMuPDF, ebooklib, beautifulsoup4, and mobi for full functionality.", "ERROR")

    def _build_ui(self):
        # Menu bar
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)
        
        file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Document Queue...", command=self.import_queue)
        file_menu.add_command(label="Export Document Queue...", command=self.export_queue)
        file_menu.add_separator()
        file_menu.add_command(label="Quit", command=self._on_closing)
        
        tools_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Check for Duplicates", command=self.check_for_duplicates)
        
        help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Help...", command=self.show_help)
        help_menu.add_command(label="About...", command=self.show_about)
        
        # Main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self._create_code_bundler_tab()
        if DOCUMENT_SUPPORT:
            self._create_document_processor_tab()
        
        # Shared log output at the bottom
        log_frame = tk.Frame(self.root)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        tk.Label(log_frame, text="Log Output:", anchor="w").pack(fill=tk.X)
        
        self.log_display = scrolledtext.ScrolledText(
            log_frame, height=12, state='disabled', wrap=tk.WORD,
            bg="#2b2b2b", fg="white", font=("Consolas", 9)
        )
        self.log_display.pack(fill=tk.BOTH, expand=True)
        
        # Configure log tags
        self.log_display.tag_config('INFO', foreground='white')
        self.log_display.tag_config('SUCCESS', foreground='#4CAF50')
        self.log_display.tag_config('ERROR', foreground='#f44336')
        self.log_display.tag_config('SUMMARY', foreground='cyan')
        self.log_display.tag_config('HEADER', foreground='yellow')
        
        # Progress bar
        self.progress = ttk.Progressbar(self.root, mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X, padx=10, pady=(0, 10))

    def _create_code_bundler_tab(self):
        """Create the project code bundler tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Project Code Bundler")
        
        # Configuration frame
        config_frame = tk.LabelFrame(tab, text="Bundle Configuration", padx=5, pady=5)
        config_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Project root selection
        root_frame = tk.Frame(config_frame)
        root_frame.pack(fill=tk.X, pady=2)
        tk.Label(root_frame, text="Project Root:", width=12, anchor='w').pack(side=tk.LEFT)
        self.project_root_var = tk.StringVar(value=os.getcwd())
        tk.Entry(root_frame, textvariable=self.project_root_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(root_frame, text="Browse", command=self._select_project_root).pack(side=tk.RIGHT)
        
        # Output file selection
        output_frame = tk.Frame(config_frame)
        output_frame.pack(fill=tk.X, pady=2)
        tk.Label(output_frame, text="Output File:", width=12, anchor='w').pack(side=tk.LEFT)
        self.bundle_output_var = tk.StringVar(value="project_bundle.txt")
        tk.Entry(output_frame, textvariable=self.bundle_output_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(output_frame, text="Browse", command=self._select_bundle_output).pack(side=tk.RIGHT)
        
        # Advanced options
        adv_frame = tk.LabelFrame(config_frame, text="Advanced Options")
        adv_frame.pack(fill=tk.X, pady=5)
        
        opts_row1 = tk.Frame(adv_frame)
        opts_row1.pack(fill=tk.X, pady=2)
        
        self.follow_symlinks_var = tk.BooleanVar()
        tk.Checkbutton(opts_row1, text="Follow Symlinks", variable=self.follow_symlinks_var).pack(side=tk.LEFT)
        
        tk.Label(opts_row1, text="Sort Mode:").pack(side=tk.LEFT, padx=(20, 5))
        self.sort_mode_var = tk.StringVar(value="path")
        sort_combo = ttk.Combobox(opts_row1, textvariable=self.sort_mode_var, values=["path", "size", "ext"], width=8)
        sort_combo.pack(side=tk.LEFT)
        
        tk.Label(opts_row1, text="ID Prefix:").pack(side=tk.LEFT, padx=(20, 5))
        self.id_prefix_var = tk.StringVar(value="F")
        tk.Entry(opts_row1, textvariable=self.id_prefix_var, width=5).pack(side=tk.LEFT)
        
        opts_row2 = tk.Frame(adv_frame)
        opts_row2.pack(fill=tk.X, pady=2)
        
        tk.Label(opts_row2, text="Max bytes per file:").pack(side=tk.LEFT)
        self.max_bytes_per_file_var = tk.StringVar(value="2000000")
        tk.Entry(opts_row2, textvariable=self.max_bytes_per_file_var, width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Label(opts_row2, text="Max total bytes:").pack(side=tk.LEFT, padx=(20, 5))
        self.max_total_bytes_var = tk.StringVar(value="50000000")
        tk.Entry(opts_row2, textvariable=self.max_total_bytes_var, width=10).pack(side=tk.LEFT, padx=5)
        
        tk.Label(opts_row2, text="LLM Guide:").pack(side=tk.LEFT, padx=(20, 5))
        self.llm_guide_var = tk.StringVar(value="short")
        guide_combo = ttk.Combobox(opts_row2, textvariable=self.llm_guide_var, values=["short", "verbose", "none"], width=8)
        guide_combo.pack(side=tk.LEFT)
        
        # Action buttons
        action_frame = tk.Frame(tab)
        action_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.bundle_button = tk.Button(
            action_frame, text="Create Project Bundle", 
            command=self.start_code_bundling, bg="#4CAF50", fg="white", font=("Arial", 10, "bold")
        )
        self.bundle_button.pack(side=tk.LEFT, padx=5)
        
        # Preview frame
        preview_frame = tk.LabelFrame(tab, text="Project Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.project_preview = scrolledtext.ScrolledText(preview_frame, height=15, state='disabled')
        self.project_preview.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        refresh_btn = tk.Button(preview_frame, text="Refresh Preview", command=self._refresh_project_preview)
        refresh_btn.pack(pady=5)

    def _create_document_processor_tab(self):
        """Create the document processor tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Document Processor")
        
        # Controls frame
        controls_frame = tk.Frame(tab)
        controls_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # File management buttons
        file_buttons = tk.Frame(controls_frame)
        file_buttons.pack(fill=tk.X, pady=5)
        
        self.add_docs_button = tk.Button(file_buttons, text="Add Documents", command=self.add_files)
        self.add_docs_button.pack(side=tk.LEFT, padx=5)
        
        self.process_docs_button = tk.Button(
            file_buttons, text="Create Knowledge Files (Batch)", 
            command=self.start_processing, bg="#4CAF50", fg="white"
        )
        self.process_docs_button.pack(side=tk.LEFT, padx=5)
        
        self.ref_sheet_button = tk.Button(
            file_buttons, text="Create Reference Sheet", 
            command=self.start_reference_sheet_creation, bg="#2196F3", fg="white"
        )
        self.ref_sheet_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_queue_button = tk.Button(
            file_buttons, text="Clear Queue", 
            command=self.clear_queue, bg="#f44336", fg="white"
        )
        self.clear_queue_button.pack(side=tk.RIGHT, padx=5)
        
        # Document queue
        queue_frame = tk.LabelFrame(tab, text="Document Queue")
        queue_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(queue_frame, text="Files to Process (will be sorted alphabetically):", anchor="w").pack(fill=tk.X, padx=5)
        
        self.queue_display = scrolledtext.ScrolledText(queue_frame, height=15, state='disabled')
        self.queue_display.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def _select_project_root(self):
        """Select project root directory."""
        directory = filedialog.askdirectory(title="Select Project Root Directory")
        if directory:
            self.project_root_var.set(directory)
            self._refresh_project_preview()
    
    def _select_bundle_output(self):
        """Select output file for project bundle."""
        filename = filedialog.asksaveasfilename(
            title="Save Project Bundle As",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if filename:
            self.bundle_output_var.set(filename)
    
    def _refresh_project_preview(self):
        """Refresh the project preview display."""
        try:
            project_root = Path(self.project_root_var.get())
            if not project_root.exists():
                self._update_preview("Invalid project root path.")
                return
                
            ignore_patterns = list(DEFAULT_IGNORE)
            ignore_patterns += read_gitignore_patterns(project_root)
            
            files = list(iter_files(project_root, ignore_patterns, self.follow_symlinks_var.get()))
            files.sort(key=lambda p: str(p))
            
            # Create a preview of the project structure
            preview_lines = [f"Project Root: {project_root}", f"Total Files: {len(files)}", "", "Project Tree:"]
            
            visible_text_files = [f for f in files if not is_probably_binary(f)]
            tree_map = make_tree_map(project_root, visible_text_files[:50])  # Limit preview
            preview_lines.append(tree_map)
            
            if len(visible_text_files) > 50:
                preview_lines.append(f"\n... and {len(visible_text_files) - 50} more files")
            
            self._update_preview("\n".join(preview_lines))
            
        except Exception as e:
            self._update_preview(f"Error generating preview: {e}")
    
    def _update_preview(self, text: str):
        """Update the project preview text."""
        self.project_preview.config(state='normal')
        self.project_preview.delete('1.0', tk.END)
        self.project_preview.insert('1.0', text)
        self.project_preview.config(state='disabled')

    # =======================================================================================
    # CODE BUNDLER FUNCTIONALITY
    # =======================================================================================
    
    def start_code_bundling(self):
        """Start the code bundling process."""
        try:
            project_root = Path(self.project_root_var.get())
            if not project_root.exists():
                messagebox.showerror("Error", "Project root directory does not exist.")
                return
            
            output_file = Path(self.bundle_output_var.get())
            
            # Validate numeric inputs
            try:
                max_bytes_per_file = int(self.max_bytes_per_file_var.get())
                max_total_bytes = int(self.max_total_bytes_var.get())
            except ValueError:
                messagebox.showerror("Error", "Max bytes values must be integers.")
                return
            
            # Start bundling in worker thread
            self._start_worker_thread(
                self._bundle_project_worker,
                (project_root, output_file, max_bytes_per_file, max_total_bytes)
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start bundling: {e}")
    
    def _bundle_project_worker(self, project_root: Path, output_file: Path, max_bytes_per_file: int, max_total_bytes: int):
        """Worker thread for project bundling."""
        try:
            self.log_queue.put(("Starting project bundling...", 'HEADER'))
            self.log_queue.put((f"Project root: {project_root}", 'INFO'))
            self.log_queue.put((f"Output file: {output_file}", 'INFO'))
            
            ignore_patterns = list(DEFAULT_IGNORE)
            ignore_patterns += read_gitignore_patterns(project_root)
            
            # Avoid bundling the output file and this script itself
            ignore_patterns.append(output_file.name)
            ignore_patterns.append(Path(__file__).name)
            
            # Collect files
            files = list(iter_files(project_root, ignore_patterns, self.follow_symlinks_var.get()))
            
            # Sort deterministically
            sort_mode = self.sort_mode_var.get()
            if sort_mode == "size":
                files.sort(key=lambda p: (p.stat().st_size if p.exists() else 0, str(p)))
            elif sort_mode == "ext":
                files.sort(key=lambda p: (p.suffix.lower(), str(p)))
            else:
                files.sort(key=lambda p: str(p))
            
            self.log_queue.put((f"Found {len(files)} files to process", 'INFO'))
            
            # Generate stable IDs
            id_prefix = self.id_prefix_var.get() or "F"
            id_width = max(4, len(str(len(files))))
            
            def file_id(i: int) -> str:
                return f"{id_prefix}{i:0{id_width}d}"
            
            # Pre-measure & pre-hash
            meta = []
            for i, p in enumerate(files, start=1):
                rel = p.relative_to(project_root).as_posix()
                lang = detect_language(p)
                is_bin = is_probably_binary(p)
                size = p.stat().st_size if p.exists() else 0
                
                sha = ""
                lines_count = 0
                note = ""
                if is_bin:
                    note = "binary: skipped"
                else:
                    try:
                        raw_full = p.read_bytes()
                    except Exception:
                        raw_full = b""
                        note = "read error: skipped"
                    if raw_full:
                        sha = hashlib.sha1(raw_full).hexdigest()
                        raw = raw_full[:max_bytes_per_file]
                        if len(raw_full) > max_bytes_per_file:
                            note = f"truncated to {max_bytes_per_file} bytes"
                        text_preview = raw.decode("utf-8", errors="ignore")
                        lines_count = text_preview.count("\n") + (1 if text_preview and not text_preview.endswith("\n") else 0)
                
                meta.append({
                    "id": file_id(i),
                    "path": rel,
                    "lang": lang if not is_bin else "Binary",
                    "size": size,
                    "lines": lines_count if not is_bin else 0,
                    "sha1": sha,
                    "is_binary": is_bin,
                    "note": note,
                })
                
                if i % 50 == 0:  # Progress update
                    progress = int((i / len(files)) * 50)  # First 50% for analysis
                    self.log_queue.put((f"__PROGRESS__{progress}", 'INFO'))
            
            # Write bundle
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with output_file.open("w", encoding="utf-8", errors="ignore") as out:
                # Header
                out.write("# PROJECT BUNDLE\n")
                out.write(f"# Generated: {timestamp}\n")
                out.write(f"# Root: {project_root}\n")
                out.write("# Format: LLM guide + project map + file index + file sections with stable IDs\n\n")
                
                # LLM usage guide
                out.write(render_llm_usage_guide(self.llm_guide_var.get()))
                
                # Project Map
                out.write("## PROJECT MAP\n")
                out.write("```\n")
                visible_text_files = [project_root / m["path"] for m in meta if not m["is_binary"]]
                out.write(make_tree_map(project_root, visible_text_files))
                out.write("\n```\n\n")
                
                # Global Index / TOC
                out.write("## FILE INDEX (Global TOC)\n")
                out.write("| ID | Path | Lang | Bytes | Lines | SHA1 | Note |\n")
                out.write("|---:|------|------:|------:|------:|------|------|\n")
                for m in meta:
                    sha_disp = (m["sha1"][:10] + "…") if m["sha1"] else ""
                    note_disp = m["note"] or ""
                    out.write(f"| {m['id']} | {m['path']} | {m['lang']} | {m['size']} | {m['lines']} | {sha_disp} | {note_disp} |\n")
                out.write("\n")
                
                # File Sections
                out.write("---\n\n")
                written_total = 0
                
                for idx, m in enumerate(meta):
                    fid = m["id"]
                    path = project_root / m["path"]
                    out.write(f"===== FILE {fid} =====\n")
                    out.write(f"PATH: {m['path']}\n")
                    out.write(f"LANG: {m['lang']}\n")
                    out.write(f"BYTES: {m['size']}\n")
                    out.write(f"LINES: {m['lines']}\n")
                    out.write(f"SHA1: {m['sha1']}\n")
                    if m["note"]:
                        out.write(f"NOTE: {m['note']}\n")
                    out.write("\n")
                    
                    if m["is_binary"]:
                        out.write("[SKIPPED] Binary content not included.\n")
                        out.write(f"----- BEGIN CONTENT {fid} -----\n")
                        out.write("[No content]\n")
                        out.write(f"----- END CONTENT {fid} -----\n\n")
                        continue
                    
                    try:
                        raw_full = path.read_bytes()
                    except Exception:
                        out.write("[SKIPPED] Could not read file as text.\n")
                        out.write(f"----- BEGIN CONTENT {fid} -----\n")
                        out.write("[No content]\n")
                        out.write(f"----- END CONTENT {fid} -----\n\n")
                        continue
                    
                    write_bytes = raw_full[:max_bytes_per_file]
                    if written_total + len(write_bytes) > max_total_bytes:
                        out.write("[SKIPPED] Total bundle size limit reached.\n")
                        out.write(f"----- BEGIN CONTENT {fid} -----\n")
                        out.write("[No content]\n")
                        out.write(f"----- END CONTENT {fid} -----\n\n")
                        continue
                    
                    text = write_bytes.decode("utf-8", errors="ignore")
                    out.write(f"----- BEGIN CONTENT {fid} -----\n")
                    out.write(text)
                    if not text.endswith("\n"):
                        out.write("\n")
                    out.write(f"----- END CONTENT {fid} -----\n\n")
                    written_total += len(write_bytes)
                    
                    # Progress update
                    progress = int(50 + ((idx + 1) / len(meta)) * 50)  # Second 50% for writing
                    self.log_queue.put((f"__PROGRESS__{progress}", 'INFO'))
                
                out.write(f"\n✅ Project bundling complete. Files: {len(meta)} | Wrote ~{written_total} bytes\n")
            
            self.log_queue.put((f"✅ Successfully created project bundle: {output_file}", 'SUCCESS'))
            self.log_queue.put((f"Total files processed: {len(meta)}", 'SUCCESS'))
            self.log_queue.put((f"Bundle size: ~{written_total:,} bytes", 'SUCCESS'))
            
        except Exception as e:
            self.log_queue.put((f"❌ Error during bundling: {e}", 'ERROR'))

    # =======================================================================================
    # DOCUMENT PROCESSOR FUNCTIONALITY
    # =======================================================================================
    
    def add_files(self):
        """Add document files to the processing queue."""
        if not DOCUMENT_SUPPORT:
            messagebox.showerror("Error", "Document processing libraries not installed.")
            return
            
        files = filedialog.askopenfilenames(
            title="Select Documents",
            filetypes=[
                ("Supported Files", "*.txt *.md *.markdown *.rst *.csv *.tsv *.log *.json *.xml *.yaml *.yml *.toml *.ini *.cfg *.conf *.sql *.tex *.rtf *.pdf *.epub *.mobi"),
                ("Text Files", "*.txt *.md *.markdown *.rst *.csv *.tsv *.log *.json *.xml *.yaml *.yml *.toml *.ini *.cfg *.conf *.sql *.tex *.rtf"),
                ("Document Files", "*.pdf *.epub *.mobi"),
                ("All files", "*.*")
            ],
        )
        if files:
            for f in files:
                if f not in self.file_queue:
                    self.file_queue.append(f)
            self.update_queue_display()
    
    def clear_queue(self):
        """Clear the document processing queue."""
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showwarning("Busy", "Cannot clear queue while processing.")
            return
        if messagebox.askyesno("Confirm", "Are you sure you want to clear the document queue?"):
            self.file_queue.clear()
            self.update_queue_display()
    
    def update_queue_display(self):
        """Update the document queue display."""
        self.queue_display.config(state='normal')
        self.queue_display.delete('1.0', tk.END)
        if not self.file_queue:
            self.queue_display.insert(tk.END, "Document queue is empty. Use 'Add Documents' to add files.")
        else:
            display_queue = sorted(self.file_queue, key=os.path.basename)
            for i, f in enumerate(display_queue, 1):
                self.queue_display.insert(tk.END, f"{i}. {os.path.basename(f)}\n")
        self.queue_display.config(state='disabled')
    
    def check_for_duplicates(self):
        """Check the document queue for duplicate filenames."""
        if not self.file_queue:
            messagebox.showinfo("Check for Duplicates", "The document queue is empty.")
            return
        
        seen = {}
        duplicates = []
        for f in self.file_queue:
            name = os.path.basename(f).lower()
            if name in seen:
                duplicates.append(f)
            else:
                seen[name] = f
        
        if not duplicates:
            messagebox.showinfo("Check for Duplicates", "No duplicates found.")
            return
        
        dup_list = "\n".join(os.path.basename(d) for d in duplicates)
        if messagebox.askyesno(
            "Duplicates Found",
            f"The following duplicates were found:\n\n{dup_list}\n\nRemove duplicates from queue?"
        ):
            self.file_queue = list(seen.values())
            self.update_queue_display()
            messagebox.showinfo("Check for Duplicates", f"Removed {len(duplicates)} duplicates from the queue.")
    
    def start_processing(self):
        """Start document processing."""
        if not DOCUMENT_SUPPORT:
            messagebox.showerror("Error", "Document processing libraries not installed.")
            return
            
        if not self.file_queue:
            messagebox.showwarning("Warning", "The document queue is empty.")
            return
        
        output_dir = filedialog.askdirectory(title="Select Output Directory for Knowledge Files")
        if not output_dir:
            return
        
        base_name = simpledialog.askstring(
            "Input", "Enter a base name for output files:", initialvalue="knowledge_base"
        )
        if not base_name:
            return
        
        chunk_size = simpledialog.askinteger(
            "Input", "How many documents per file?", initialvalue=10, minvalue=1
        )
        if not chunk_size:
            return
        
        id_prefix = simpledialog.askstring(
            "Optional", "DocID prefix (letters/numbers, default 'DOC'):", initialvalue=self.docid_prefix
        )
        if id_prefix:
            id_prefix = re.sub(r"[^A-Za-z0-9]", "", id_prefix).upper()
            if not id_prefix:
                id_prefix = "DOC"
        else:
            id_prefix = "DOC"
        self.docid_prefix = id_prefix
        
        self._start_worker_thread(
            self._process_documents_worker,
            (self.file_queue.copy(), output_dir, base_name, chunk_size, id_prefix)
        )
    
    def _process_documents_worker(self, file_paths: List[str], output_dir: str, base_name: str, chunk_size: int, id_prefix: str):
        """Worker thread for document processing."""
        self.log_queue.put(("Sorting document queue alphabetically...", 'INFO'))
        file_paths.sort(key=os.path.basename)
        
        total_processed, total_failed, batch_num = 0, 0, 1
        file_chunks = [file_paths[i:i + chunk_size] for i in range(0, len(file_paths), chunk_size)]
        
        total = len(file_paths)
        done = 0
        
        for chunk in file_chunks:
            output_filename = f"{base_name}_{batch_num}.txt"
            output_filepath = os.path.join(output_dir, output_filename)
            self.log_queue.put((f"\n--- Starting Batch {batch_num} -> {output_filename} ---", 'HEADER'))
            processed_docs: List[Dict[str, Any]] = []
            
            max_workers = min(max(2, (os.cpu_count() or 2)), 4)
            with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {executor.submit(process_single_file, fp, id_prefix): fp for fp in chunk}
                for future in concurrent.futures.as_completed(future_to_file):
                    src_path = future_to_file[future]
                    file_path_for_log = os.path.basename(src_path)
                    try:
                        _, result = future.result()
                    except Exception as e:
                        result = {"error": f"Worker crashed: {e}"}
                    
                    if "error" in result:
                        total_failed += 1
                        self.log_queue.put((f"  ├─ Error on '{file_path_for_log}': {result['error']}", 'ERROR'))
                    else:
                        result["_order"] = chunk.index(src_path)
                        processed_docs.append(result)
                        self.log_queue.put((f"  ├─ Success! ID: {result['short_id']} for '{file_path_for_log}'", 'SUCCESS'))
                    
                    done += 1
                    self.log_queue.put((f"__PROGRESS__{int(done * 100 / total)}", 'INFO'))
            
            if processed_docs:
                processed_docs.sort(key=lambda x: x['_order'])
                try:
                    with open(output_filepath, 'w', encoding='utf-8', newline='\n') as outfile:
                        outfile.write(KNOWLEDGE_FILE_HEADER)
                        outfile.write("\n--- TABLE OF CONTENTS ---\n")
                        for doc in processed_docs:
                            outfile.write(f"[DocID: {doc['short_id']} ({doc['full_hash']}) | Title: {doc['title']}]\n")
                        outfile.write("--- END OF TOC ---\n\n")
                        
                        for doc in processed_docs:
                            outfile.write(f"[START OF DOCUMENT: {doc['short_id']} | Title: {doc['title']}]\n\n")
                            clean = sanitize_text(doc['text'])
                            # Soft-wrap long lines
                            clean = re.sub(r'[^\n]{10000,}',
                                         lambda m: '\n'.join(m.group(0)[i:i+10000] for i in range(0, len(m.group(0)), 10000)),
                                         clean)
                            outfile.write(clean)
                            outfile.write(f"\n\n[END OF DOCUMENT: {doc['short_id']}]\n---\n\n")
                    
                    self.log_queue.put((f"✅ Batch {batch_num} complete. Wrote {len(processed_docs)} documents.", 'SUCCESS'))
                    total_processed += len(processed_docs)
                except IOError as e:
                    self.log_queue.put((f"FATAL I/O ERROR: {e}", 'ERROR'))
            else:
                self.log_queue.put((f"⚠ Batch {batch_num} had no documents to write.", 'ERROR'))
            
            batch_num += 1
        
        self.log_queue.put(("\n--- Overall Processing Complete ---", 'SUMMARY'))
        self.log_queue.put((f"Total documents processed: {total_processed}", 'SUCCESS'))
        self.log_queue.put((f"Total files failed: {total_failed}", 'ERROR'))
    
    def start_reference_sheet_creation(self):
        """Start reference sheet creation."""
        if not DOCUMENT_SUPPORT:
            messagebox.showerror("Error", "Document processing libraries not installed.")
            return
            
        input_paths = filedialog.askopenfilenames(
            title="Select Knowledge Files to Index", filetypes=[("Text Files", "*.txt")]
        )
        if not input_paths:
            return
        output_path = filedialog.asksaveasfilename(
            title="Save Reference Sheet As", defaultextension=".txt", filetypes=[("Text Files", "*.txt")]
        )
        if not output_path:
            return
        self._start_worker_thread(self._create_reference_sheet_worker, (input_paths, output_path))
    
    def _create_reference_sheet_worker(self, input_paths: List[str], output_path: str):
        """Worker thread for reference sheet creation."""
        self.log_queue.put(("Starting Reference Sheet creation...", 'SUMMARY'))
        doc_references: Dict[str, Dict[str, str]] = {}
        
        doc_id_regex = re.compile(r"^\[DocID: ([A-Z0-9]+) \((sha256-[a-f0-9]{64})\) \| Title: ([^\]]+)\]\s*$")
        
        total_matches = 0
        for path in input_paths:
            source_filename = os.path.basename(path)
            self.log_queue.put((f"Scanning: {source_filename}", 'INFO'))
            matches_in_file = 0
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as infile:
                    for line in infile:
                        m = doc_id_regex.match(line)
                        if m:
                            short_id, _, title = m.groups()
                            doc_references.setdefault(short_id, {'title': title, 'source': source_filename})
                            matches_in_file += 1
                        else:
                            if "[DocID:" in line:
                                self.log_queue.put((f"  ├─ Skipped line (format mismatch): {line.strip()}", 'INFO'))
            except Exception as e:
                self.log_queue.put((f"  ├─ Could not parse file: {e}", 'ERROR'))
            
            total_matches += matches_in_file
            self.log_queue.put((f"  ├─ Found {matches_in_file} entries", 'INFO'))
        
        if not doc_references:
            self.log_queue.put(("\nNo valid DocID entries found. Check file format and regex.", 'ERROR'))
            return
        
        sorted_refs = sorted(doc_references.items(), key=lambda item: item[1]['title'].lower())
        try:
            with open(output_path, 'w', encoding='utf-8', newline='\n') as outfile:
                outfile.write(REFERENCE_SHEET_HEADER)
                outfile.write("\n--- GLOBAL DOCUMENT INDEX ---\n")
                outfile.write("\n".join(
                    f"[DocID: {s_id} | Title: {d['title']}] | [SourceFile: {d['source']}]" for s_id, d in sorted_refs
                ))
                outfile.write("\n--- END OF INDEX ---\n")
            self.log_queue.put(("\n--- Finalization ---", 'HEADER'))
            self.log_queue.put((f"Successfully indexed {len(sorted_refs)} unique documents (from {total_matches} matches).", 'SUCCESS'))
        except IOError as e:
            self.log_queue.put((f"FATAL I/O ERROR: {e}", 'ERROR'))

    # =======================================================================================
    # QUEUE MANAGEMENT
    # =======================================================================================
    
    def import_queue(self):
        """Import document queue from file."""
        path = filedialog.askopenfilename(
            title="Import Queue File",
            filetypes=[("Queue Files", "*.que.txt"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                new_files = [line.strip() for line in f if line.strip()]
            added_count = 0
            for file_path in new_files:
                if os.path.exists(file_path) and file_path not in self.file_queue:
                    self.file_queue.append(file_path)
                    added_count += 1
            self.update_queue_display()
            messagebox.showinfo("Success", f"Imported {len(new_files)} paths.\nAdded {added_count} new, valid files.")
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not import queue file: {e}")
    
    def export_queue(self):
        """Export document queue to file."""
        if not self.file_queue:
            messagebox.showwarning("Warning", "Document queue is empty.")
            return
        path = filedialog.asksaveasfilename(
            title="Export Queue File",
            defaultextension=".que.txt",
            filetypes=[("Queue Files", "*.que.txt")],
        )
        if not path:
            return
        try:
            sorted_queue = sorted(self.file_queue, key=os.path.basename)
            with open(path, 'w', encoding='utf-8', newline='\n') as f:
                f.write("\n".join(sorted_queue))
            messagebox.showinfo("Success", f"Successfully exported {len(sorted_queue)} file paths.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export queue file: {e}")

    # =======================================================================================
    # SHARED UI FUNCTIONALITY
    # =======================================================================================
    
    def show_help(self):
        """Show help dialog."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Help")
        help_window.geometry("800x600")
        
        help_text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        help_text.pack(fill=tk.BOTH, expand=True)
        help_text.insert('1.0', HELP_TEXT)
        help_text.config(state='disabled')
        
        close_btn = tk.Button(help_window, text="Close", command=help_window.destroy)
        close_btn.pack(pady=10)
    
    def show_about(self):
        """Show about dialog."""
        about_window = tk.Toplevel(self.root)
        about_window.title("About")
        about_window.geometry("600x400")
        about_window.resizable(False, False)
        
        # Center the window
        about_window.transient(self.root)
        about_window.grab_set()
        
        main_frame = tk.Frame(about_window, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(main_frame, text="Unified LLM Content Preparation Tool", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Version
        version_label = tk.Label(main_frame, text="Version 1.0", font=("Arial", 12))
        version_label.pack()
        
        # About text
        about_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=15, width=60)
        about_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        about_text.insert('1.0', ABOUT_TEXT)
        about_text.config(state='disabled')
        
        # Close button
        close_btn = tk.Button(main_frame, text="Close", command=about_window.destroy, width=10)
        close_btn.pack(pady=(10, 0))
    
    def log_message(self, message: str, level: str = 'INFO'):
        """Add a message to the log display."""
        # Handle progress updates specially
        if message.startswith("__PROGRESS__"):
            try:
                pct = int(message.split("__PROGRESS__")[1])
                self.progress['value'] = pct
                self.root.update_idletasks()
                return
            except Exception:
                pass
        
        self.log_display.config(state='normal')
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_display.insert(tk.END, f"[{timestamp}] {message}\n", level.upper())
        self.log_display.config(state='disabled')
        self.log_display.see(tk.END)
        self.root.update_idletasks()
    
    def _start_worker_thread(self, target_func, args_tuple):
        """Start a worker thread and manage UI state."""
        self._set_ui_state(False)
        
        # Clear log
        self.log_display.config(state='normal')
        self.log_display.delete('1.0', tk.END)
        self.log_display.config(state='disabled')
        self.progress['value'] = 0
        
        # Start worker thread
        self.processing_thread = threading.Thread(target=target_func, args=args_tuple, daemon=True)
        self.processing_thread.start()
        self.root.after(100, self._check_log_queue)
    
    def _set_ui_state(self, enabled: bool):
        """Enable or disable UI controls during processing."""
        state = 'normal' if enabled else 'disabled'
        
        # Code bundler controls
        self.bundle_button.config(state=state)
        
        # Document processor controls
        if DOCUMENT_SUPPORT:
            self.add_docs_button.config(state=state)
            self.process_docs_button.config(state=state)
            self.ref_sheet_button.config(state=state)
            self.clear_queue_button.config(state=state)
    
    def _check_log_queue(self):
        """Check for log messages from worker threads."""
        try:
            while True:
                message, level = self.log_queue.get_nowait()
                self.log_message(message, level)
        except queue.Empty:
            pass
        
        if self.processing_thread and self.processing_thread.is_alive():
            self.root.after(100, self._check_log_queue)
        else:
            self._set_ui_state(True)
            self.progress['value'] = 0
    
    def _on_closing(self):
        """Handle application closing."""
        if self.processing_thread and self.processing_thread.is_alive():
            if messagebox.askyesno("Exit", "Processing is active. Are you sure you want to exit?"):
                self.root.destroy()
        else:
            self.root.destroy()


# =======================================================================================
# MAIN APPLICATION ENTRY POINT
# =======================================================================================

def main():
    """Main application entry point."""
    # Ensure multiprocessing works on Windows
    multiprocessing.freeze_support()
    
    # Create and run the application
    root = tk.Tk()
    
    # Set application icon if available
    try:
        # You can add an icon file here if desired
        # root.iconbitmap('icon.ico')
        pass
    except Exception:
        pass
    
    app = UnifiedLLMPrepTool(root)
    
    # Center the window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    # Start the main loop
    root.mainloop()


if __name__ == "__main__":
    main()