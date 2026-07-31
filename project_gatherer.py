#!/usr/bin/env python3
"""
Project Gatherer
================

A portable, cross-platform (macOS / Windows / Linux) tool that lives on an
external drive. Point it at one or more drives and it will scan them for every
image and video that has even the slightest chance of being associated with one
of your projects, then copy the matches into a uniquely-named folder on the
drive the app lives on.

How it decides what is "project related"
----------------------------------------
You supply a Markdown file (typically your transcribed descriptions of every
project you can remember). The app reads that file and extracts distinctive
"project terms": names in Markdown headings, bold or quoted text, capitalised
proper-noun phrases, years, codes and rare words. Common English words and
filler/transcription words are ignored.

Every candidate file is scored by matching those terms against BOTH its file
name and its full folder path (so a generically-named IMG_2931.jpg sitting in a
".../Clients/Ritual Industries/shoot/" folder still gets caught). PowerPoint
files (.pptx / .pptm / .ppsx / .potx) are opened as zip archives: their slide
text is matched too, and when a deck matches, the images and videos embedded in
its ppt/media/ folder are extracted.

Safety
------
The scanned drives are only ever READ. Nothing on them is moved, deleted or
modified. All output is written to a new folder on the app's own drive.

Requires: Python 3.8+ (Tkinter ships with the standard python.org installers on
both macOS and Windows). No third-party packages.
"""

import csv
import os
import queue
import re
import shutil
import socket
import string
import sys
import threading
import time
import traceback
import zipfile
from datetime import datetime

# ---------------------------------------------------------------------------
# File type knowledge
# ---------------------------------------------------------------------------

IMAGE_EXTS = {
    # common raster
    "jpg", "jpeg", "jpe", "jfif", "png", "gif", "bmp", "dib", "tif", "tiff",
    "webp", "heic", "heif", "avif", "ico",
    # camera raw
    "raw", "cr2", "cr3", "crw", "nef", "nrw", "arw", "srf", "sr2", "dng",
    "orf", "rw2", "raf", "srw", "pef", "x3f", "3fr", "erf", "kdc", "mos",
    "iiq", "rwl", "fff",
    # design / layered
    "psd", "psb", "ai", "eps", "svg", "indd", "sketch", "xcf", "cr2",
}

VIDEO_EXTS = {
    "mp4", "m4v", "mov", "qt", "avi", "mkv", "wmv", "flv", "f4v", "webm",
    "mpg", "mpeg", "mpe", "m2v", "mp2", "3gp", "3g2", "mts", "m2ts", "ts",
    "vob", "ogv", "rm", "rmvb", "asf", "divx", "mxf", "r3d", "braw", "prores",
    "dv", "m4p",
}

# PowerPoint files that are actually zip archives.
PPTX_EXTS = {"pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "thmx"}

MEDIA_EXTS = IMAGE_EXTS | VIDEO_EXTS

# Directories we never descend into: OS/system noise, caches, app internals.
# Skipping these dramatically speeds scans and avoids false positives while
# leaving all normal user document / photo / project locations untouched.
SKIP_DIR_NAMES = {
    # cross platform junk
    ".git", ".svn", ".hg", "node_modules", "__pycache__", ".cache", ".Trash",
    ".Trashes", ".Spotlight-V100", ".fseventsd", ".DocumentRevisions-V100",
    ".TemporaryItems", ".DS_Store", "$RECYCLE.BIN", "$Recycle.Bin",
    "System Volume Information", "found.000",
    # macOS system
    "System", "private", "Library", "cores", "dev", "Network",
    # Windows system
    "Windows", "Program Files", "Program Files (x86)", "ProgramData",
    "$WinREAgent", "Recovery", "PerfLogs", "AppData",
    # Linux system
    "proc", "sys", "run", "bin", "sbin", "lib", "lib64", "usr", "boot",
}

# Our own output folders always start with this prefix; never rescan them.
OUTPUT_PREFIX = "ProjectGather"

APP_NAME = "Project Gatherer"

# ---------------------------------------------------------------------------
# Text normalisation & term extraction
# ---------------------------------------------------------------------------

_CAMEL_1 = re.compile(r"([a-z0-9])([A-Z])")
_CAMEL_2 = re.compile(r"([A-Z]+)([A-Z][a-z])")
_NON_ALNUM = re.compile(r"[^0-9a-z]+")


def normalize(text):
    """Return (token_list, joined_string) for a chunk of text.

    Splits camelCase / PascalCase and treats every non-alphanumeric character
    as a separator, so "Ritual-Industries_FINAL.v2" and "ritual industries
    final v2" tokenize identically.
    """
    if not text:
        return [], ""
    text = _CAMEL_1.sub(r"\1 \2", text)
    text = _CAMEL_2.sub(r"\1 \2", text)
    text = text.lower()
    text = _NON_ALNUM.sub(" ", text)
    tokens = [t for t in text.split() if t]
    return tokens, " ".join(tokens)


# Words that must never become project terms on their own. Standard English
# stopwords plus filler words common in spoken transcriptions and generic
# media/production vocabulary that would over-match if used alone.
STOPWORDS = set("""
a about above after again against all am an and any are aren't as at be
because been before being below between both but by can can't cannot could
couldn't did didn't do does doesn't doing don't down during each few for from
further had hadn't has hasn't have haven't having he he'd he'll he's her here
here's hers herself him himself his how how's i i'd i'll i'm i've if in into is
isn't it it's its itself let's me more most mustn't my myself no nor not of off
on once only or other ought our ours ourselves out over own same shan't she
she'd she'll she's should shouldn't so some such than that that's the their
theirs them themselves then there there's these they they'd they'll they're
they've this those through to too under until up very was wasn't we we'd we'll
we're we've were weren't what what's when when's where where's which while who
who's whom why why's with won't would wouldn't you you'd you'll you're you've
your yours yourself yourselves
um uh eh ah oh okay ok yeah yep nope like just really actually basically
literally kind sort thing things stuff gonna wanna gotta maybe probably guess
mean know think thought remember remembered talking talked said say saying
went get got getting one two three first second next also well right sure lot
lots bit little big small good great nice cool done doing did make made making
went come came going go back then now today day time year years ago around
something anything everything someone anyone everyone somewhere anywhere
project projects work worked working file files folder photo photos photograph
image images picture pictures video videos footage clip clips shoot shot shots
final draft version media asset assets content client clients studio design
designs edit edited editing render rendered scene camera team people company
""".split())

# Capitalised words that are too generic to trigger on their own even when they
# appear as proper nouns in the projects file.
GENERIC_CAPS = {
    "project", "projects", "studio", "design", "media", "video", "photo",
    "photos", "final", "draft", "client", "shoot", "camera", "team", "company",
    "the", "and", "of", "for", "with", "inc", "llc", "co", "group",
}


def _phrase_key(tokens):
    return " ".join(tokens)


class ProjectProfile:
    """Extracts weighted project terms from the projects Markdown file and
    scores arbitrary file paths against them."""

    # Weights by evidence strength.
    W_HEADER = 5.0        # Markdown heading -> almost certainly a project name
    W_EMPH = 4.0          # bold / quoted text -> deliberately emphasised
    W_PHRASE = 3.0        # multi-word capitalised proper noun
    W_PROPER = 2.0        # single capitalised proper noun
    W_RARE = 1.2          # rare lowercase word
    W_CODE = 1.5          # alphanumeric code (e.g. PRJ2023)
    W_SUPPORT = 0.5       # years / generic caps -> supporting evidence only

    def __init__(self):
        # Multi-word phrases: list of dicts {key, tokens, weight, display}.
        self.phrases = []
        # Single trigger tokens: token -> weight.
        self.tokens = {}
        # Supporting tokens (years, generic caps): never trigger on their own.
        self.support = {}
        # Fuzzy index (SymSpell-style deletes) built lazily when enabled.
        self._delete_index = None
        self._fuzzy_tokens = {}

    # -- construction -------------------------------------------------------

    def _add_token(self, tok, weight):
        if not tok or tok in STOPWORDS or len(tok) < 3:
            return
        self.tokens[tok] = max(self.tokens.get(tok, 0.0), weight)

    def _add_support(self, tok, weight=W_SUPPORT):
        if tok:
            self.support[tok] = max(self.support.get(tok, 0.0), weight)

    def _add_phrase(self, text, weight):
        tokens, joined = normalize(text)
        # Drop stopword-only tokens from the phrase.
        useful = [t for t in tokens if t not in STOPWORDS and len(t) >= 2]
        if not useful:
            return
        if len(useful) >= 2:
            key = _phrase_key(useful)
            if not any(p["key"] == key for p in self.phrases):
                self.phrases.append({
                    "key": key,
                    "tokens": useful,
                    "weight": weight,
                    "display": text.strip(),
                })
            # Also index the strongest single tokens of a phrase so a file
            # named after just one word of the project still matches (weaker).
            for t in useful:
                if t not in GENERIC_CAPS:
                    self._add_token(t, max(self.W_PROPER, weight - 2.0))
        else:
            t = useful[0]
            if t in GENERIC_CAPS:
                self._add_support(t)
            else:
                self._add_token(t, weight)

    def build_from_markdown(self, md_text):
        # Word-frequency table to find genuinely rare/distinctive lowercase
        # words (rare in the description == distinctive == good signal).
        all_tokens, _ = normalize(md_text)
        freq = {}
        for t in all_tokens:
            freq[t] = freq.get(t, 0) + 1

        for raw_line in md_text.splitlines():
            line = raw_line.rstrip()

            # Markdown headings -> project names (strongest signal).
            m = re.match(r"^\s{0,3}#{1,6}\s+(.*\S)", line)
            if m:
                self._add_phrase(m.group(1), self.W_HEADER)

            # Bullet leaders that often precede a project title.
            m2 = re.match(r"^\s*[-*+]\s+(.*\S)", line)
            if m2 and re.match(r"^[A-Z0-9\"'“]", m2.group(1)):
                # Take the leading capitalised run of the bullet.
                lead = re.match(r"([A-Z0-9][\w&'’.\- ]{0,60})", m2.group(1))
                if lead:
                    self._add_phrase(lead.group(1), self.W_PHRASE)

        # Bold / italic emphasis.
        for m in re.finditer(r"\*\*(.+?)\*\*|__(.+?)__|\*(.+?)\*|_(.+?)_", md_text):
            text = next(g for g in m.groups() if g)
            self._add_phrase(text, self.W_EMPH)

        # Quoted text (straight and curly quotes).
        for m in re.finditer(r"[\"“”]([^\"“”\n]{2,80})[\"“”]"
                             r"|[‘’']([^‘’'\n]{2,80})[‘’']",
                             md_text):
            text = next(g for g in m.groups() if g)
            self._add_phrase(text, self.W_EMPH)

        # Capitalised proper-noun phrases anywhere in the body.
        # A run of Capitalised words, optionally joined by small connectors.
        cap_run = re.compile(
            r"\b([A-Z][A-Za-z0-9&'’]+(?:\s+(?:of|the|and|for|de|la|los|las|del|&|"
            r"[A-Z][A-Za-z0-9&'’]+))*)"
        )
        for m in cap_run.finditer(md_text):
            phrase = m.group(1).strip()
            words = phrase.split()
            if len(words) >= 2:
                self._add_phrase(phrase, self.W_PHRASE)
            else:
                low = words[0].lower()
                if low in GENERIC_CAPS:
                    self._add_support(low)
                elif low not in STOPWORDS and len(low) >= 3:
                    self._add_token(low, self.W_PROPER)

        # Years (supporting evidence only -> never trigger alone).
        for m in re.finditer(r"\b(19\d{2}|20\d{2})\b", md_text):
            self._add_support(m.group(1))

        # Alphanumeric codes such as PRJ2023, v2rev3, SKU-4451.
        for m in re.finditer(r"\b(?=[a-z0-9]*\d)(?=[a-z0-9]*[a-z])[a-z0-9]{4,}\b",
                             md_text, re.IGNORECASE):
            code = m.group(0).lower()
            if code not in STOPWORDS:
                self._add_token(code, self.W_CODE)

        # Rare distinctive lowercase words (appear only once or twice, long).
        for tok, count in freq.items():
            if (count <= 2 and len(tok) >= 7 and tok not in STOPWORDS
                    and tok not in self.tokens):
                self._add_token(tok, self.W_RARE)

        return self

    def term_count(self):
        return len(self.phrases) + len(self.tokens)

    def summary(self):
        strong = sorted(self.phrases, key=lambda p: -p["weight"])[:12]
        names = [p["display"] for p in strong]
        return names

    # -- fuzzy index --------------------------------------------------------

    @staticmethod
    def _deletes(word):
        """1-edit deletion variants of a word (SymSpell trick)."""
        return {word[:i] + word[i + 1:] for i in range(len(word))}

    def _build_fuzzy(self):
        index = {}
        fuzzy_tokens = {}
        for tok, w in self.tokens.items():
            if len(tok) >= 6:
                fuzzy_tokens[tok] = w
                index.setdefault(tok, set()).add(tok)
                for d in self._deletes(tok):
                    index.setdefault(d, set()).add(tok)
        self._delete_index = index
        self._fuzzy_tokens = fuzzy_tokens

    @staticmethod
    def _within_edit1(a, b):
        if a == b:
            return True
        la, lb = len(a), len(b)
        if abs(la - lb) > 1:
            return False
        if la == lb:  # substitution
            return sum(1 for x, y in zip(a, b) if x != y) == 1
        # insertion / deletion
        if la > lb:
            a, b = b, a
            la, lb = lb, la
        i = j = 0
        diff = 0
        while i < la and j < lb:
            if a[i] != b[j]:
                diff += 1
                if diff > 1:
                    return False
                j += 1
            else:
                i += 1
                j += 1
        return True

    # -- scoring ------------------------------------------------------------

    def score(self, tokens, joined, fuzzy=False, extra_tokens=None):
        """Score a candidate. `tokens` is a set of normalized tokens from the
        file name + folder path; `joined` is the space-joined normalized path
        used for phrase substring matching. Returns (score, reasons)."""
        tokset = tokens if isinstance(tokens, set) else set(tokens)
        if extra_tokens:
            tokset = tokset | extra_tokens
        primary = 0.0
        reasons = []

        # Multi-word phrases: full weight if the phrase appears as a substring
        # of the path, 80% if all its tokens are present but scattered.
        for p in self.phrases:
            if p["key"] in joined:
                primary += p["weight"]
                reasons.append(f"phrase:{p['display']}")
            elif all(t in tokset for t in p["tokens"]):
                primary += p["weight"] * 0.8
                reasons.append(f"phrase~:{p['display']}")

        # Single trigger tokens (exact).
        matched_tokens = []
        for tok, w in self.tokens.items():
            if tok in tokset:
                primary += w
                matched_tokens.append(tok)
        if matched_tokens:
            reasons.append("terms:" + ",".join(sorted(matched_tokens)[:8]))

        # Optional fuzzy matching for typos / minor variants.
        if fuzzy:
            if self._delete_index is None:
                self._build_fuzzy()
            fuzzy_hits = []
            for ft in tokset:
                if len(ft) < 6 or ft in self.tokens:
                    continue
                candidates = set()
                if ft in self._delete_index:
                    candidates |= self._delete_index[ft]
                for d in self._deletes(ft):
                    if d in self._delete_index:
                        candidates |= self._delete_index[d]
                for cand in candidates:
                    if cand not in matched_tokens and self._within_edit1(ft, cand):
                        primary += self._fuzzy_tokens[cand] * 0.5
                        fuzzy_hits.append(f"{ft}~{cand}")
                        break
            if fuzzy_hits:
                reasons.append("fuzzy:" + ",".join(fuzzy_hits[:6]))

        # Supporting evidence (years, generic caps) only counts if there is
        # already at least some real signal.
        if primary > 0:
            support_hits = [s for s in self.support if s in tokset]
            if support_hits:
                primary += sum(self.support[s] for s in support_hits)
                reasons.append("support:" + ",".join(sorted(support_hits)[:6]))

        return primary, reasons


# Sensitivity -> minimum score to keep a file.
SENSITIVITY = {
    "High recall (catch almost anything)": 0.9,
    "Balanced": 2.0,
    "Strict (high confidence only)": 4.0,
}


# ---------------------------------------------------------------------------
# Drive discovery
# ---------------------------------------------------------------------------

def human_size(n):
    for unit in ("B", "KB", "MB", "GB", "TB", "PB"):
        if n < 1024:
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} EB"


def _free_bytes(path):
    try:
        return shutil.disk_usage(path).free
    except Exception:
        return None


def list_drives():
    """Return a list of (label, path) tuples for user-selectable drives."""
    drives = []
    if sys.platform.startswith("win"):
        try:
            import ctypes
            for letter in string.ascii_uppercase:
                root = f"{letter}:\\"
                if os.path.exists(root):
                    name = ""
                    try:
                        buf = ctypes.create_unicode_buffer(1024)
                        ctypes.windll.kernel32.GetVolumeInformationW(
                            ctypes.c_wchar_p(root), buf, ctypes.sizeof(buf),
                            None, None, None, None, 0)
                        name = buf.value
                    except Exception:
                        pass
                    free = _free_bytes(root)
                    label = f"{letter}:\\  {name}".rstrip()
                    if free is not None:
                        label += f"   ({human_size(free)} free)"
                    drives.append((label, root))
        except Exception:
            for letter in string.ascii_uppercase:
                root = f"{letter}:\\"
                if os.path.exists(root):
                    drives.append((root, root))
    elif sys.platform == "darwin":
        vol = "/Volumes"
        seen = set()
        if os.path.isdir(vol):
            for name in sorted(os.listdir(vol)):
                p = os.path.join(vol, name)
                if os.path.isdir(p) and not name.startswith("."):
                    free = _free_bytes(p)
                    label = name + (f"   ({human_size(free)} free)" if free else "")
                    drives.append((label, p))
                    seen.add(os.path.realpath(p))
        # Boot volume, if not already covered by a /Volumes symlink.
        if os.path.realpath("/") not in seen:
            free = _free_bytes("/")
            drives.insert(0, ("Macintosh HD (/)" +
                              (f"   ({human_size(free)} free)" if free else ""), "/"))
    else:  # linux and friends
        for base in ("/media", "/mnt", "/run/media"):
            if os.path.isdir(base):
                for root, dirs, _files in os.walk(base):
                    for d in dirs:
                        p = os.path.join(root, d)
                        if os.path.ismount(p):
                            free = _free_bytes(p)
                            label = p + (f"   ({human_size(free)} free)" if free else "")
                            drives.append((label, p))
                    break
        free = _free_bytes("/")
        drives.append(("/ (root filesystem)" +
                       (f"   ({human_size(free)} free)" if free else ""), "/"))
    return drives


def app_drive_dir():
    """Directory the app lives in (output is created next to it)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def sanitize(name):
    name = re.sub(r"[^\w.\- ]+", "_", name).strip(" .")
    return name or "unnamed"


def volume_folder_name(root):
    """A stable, filesystem-safe name for a scanned drive."""
    if sys.platform.startswith("win"):
        letter = root.rstrip(":\\/") or root
        return sanitize(f"{letter}_drive")
    if root == "/":
        return "root"
    return sanitize(os.path.basename(root.rstrip("/")) or "root")


# ---------------------------------------------------------------------------
# Scanner core
# ---------------------------------------------------------------------------

class Stop(Exception):
    pass


class Scanner:
    def __init__(self, profile, roots, output_dir, sensitivity, fuzzy,
                 do_pptx, copy_enabled, msg_queue, stop_event):
        self.profile = profile
        self.roots = roots
        self.output_dir = output_dir
        self.threshold = sensitivity
        self.fuzzy = fuzzy
        self.do_pptx = do_pptx
        self.copy_enabled = copy_enabled
        self.q = msg_queue
        self.stop = stop_event

        self.output_real = os.path.realpath(output_dir)
        self.stats = {
            "scanned": 0, "media_matched": 0, "pptx_scanned": 0,
            "pptx_matched": 0, "pptx_media": 0, "copied": 0,
            "bytes": 0, "errors": 0,
        }
        self.manifest_rows = []

    # -- messaging ----------------------------------------------------------

    def log(self, text):
        self.q.put(("log", text))

    def progress(self):
        self.q.put(("progress", dict(self.stats)))

    def check_stop(self):
        if self.stop.is_set():
            raise Stop()

    # -- path helpers -------------------------------------------------------

    def _path_signals(self, root, full_path):
        """Return (tokset, joined) built from the file name + every folder
        component between the drive root and the file."""
        rel = os.path.relpath(full_path, root)
        parts = re.split(r"[\\/]+", rel)
        name = parts[-1]
        stem = os.path.splitext(name)[0]
        pieces = parts[:-1] + [stem]
        tokset = set()
        joined_parts = []
        for piece in pieces:
            toks, joined = normalize(piece)
            tokset.update(toks)
            if joined:
                joined_parts.append(joined)
        return tokset, " ".join(joined_parts)

    def _dest_for(self, root, full_path):
        vol = volume_folder_name(root)
        rel = os.path.relpath(full_path, root)
        dest = os.path.join(self.output_dir, vol, rel)
        return dest

    def _copy(self, src, dest, category, size, mtime, score, reasons, source_note=""):
        if self.copy_enabled:
            try:
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                if not os.path.exists(dest):
                    shutil.copy2(src, dest)
                    self.stats["copied"] += 1
                    self.stats["bytes"] += size
            except Exception as e:
                self.stats["errors"] += 1
                self.log(f"  ! copy failed: {src} ({e})")
                return
        self.manifest_rows.append({
            "category": category,
            "source": source_note or src,
            "dest": os.path.relpath(dest, self.output_dir) if self.copy_enabled else "(scan only)",
            "size_bytes": size,
            "modified": datetime.fromtimestamp(mtime).isoformat(timespec="seconds") if mtime else "",
            "score": f"{score:.2f}",
            "reasons": " | ".join(reasons),
        })

    # -- pptx ---------------------------------------------------------------

    _XML_TAG = re.compile(r"<[^>]+>")

    def _handle_pptx(self, root, full_path):
        self.stats["pptx_scanned"] += 1
        try:
            with zipfile.ZipFile(full_path) as z:
                names = z.namelist()
                media = [n for n in names
                         if n.lower().startswith("ppt/media/")
                         and os.path.splitext(n)[1].lstrip(".").lower() in MEDIA_EXTS]
                if not media:
                    return
                # Build signals from path + slide/notes text for matching.
                tokset, joined = self._path_signals(root, full_path)
                text_tokens = set()
                for n in names:
                    ln = n.lower()
                    if (ln.startswith("ppt/slides/slide") or
                            ln.startswith("ppt/notesslides/")) and ln.endswith(".xml"):
                        try:
                            raw = z.read(n).decode("utf-8", "ignore")
                        except Exception:
                            continue
                        text = self._XML_TAG.sub(" ", raw)
                        toks, _ = normalize(text)
                        text_tokens.update(toks)
                score, reasons = self.profile.score(
                    tokset, joined, fuzzy=self.fuzzy, extra_tokens=text_tokens)
                if score < self.threshold:
                    return
                self.stats["pptx_matched"] += 1
                reasons = ["(embedded in PPT)"] + reasons
                deck_rel = os.path.relpath(full_path, root)
                deck_stem = os.path.splitext(os.path.basename(full_path))[0]
                vol = volume_folder_name(root)
                out_base = os.path.join(
                    self.output_dir, vol,
                    os.path.dirname(deck_rel),
                    sanitize(deck_stem) + "__embedded_media")
                for n in media:
                    self.check_stop()
                    mname = os.path.basename(n)
                    dest = os.path.join(out_base, mname)
                    try:
                        info = z.getinfo(n)
                        size = info.file_size
                    except Exception:
                        size = 0
                    if self.copy_enabled:
                        try:
                            os.makedirs(out_base, exist_ok=True)
                            if not os.path.exists(dest):
                                with z.open(n) as fsrc, open(dest, "wb") as fdst:
                                    shutil.copyfileobj(fsrc, fdst)
                                self.stats["copied"] += 1
                                self.stats["pptx_media"] += 1
                                self.stats["bytes"] += size
                        except Exception as e:
                            self.stats["errors"] += 1
                            self.log(f"  ! extract failed: {n} in {full_path} ({e})")
                            continue
                    else:
                        self.stats["pptx_media"] += 1
                    self.manifest_rows.append({
                        "category": "pptx-media",
                        "source": f"{full_path} :: {n}",
                        "dest": os.path.relpath(dest, self.output_dir) if self.copy_enabled else "(scan only)",
                        "size_bytes": size,
                        "modified": "",
                        "score": f"{score:.2f}",
                        "reasons": " | ".join(reasons),
                    })
                self.log(f"  + PPT match ({score:.1f}): {deck_rel}  [{len(media)} media]")
        except zipfile.BadZipFile:
            pass  # legacy .ppt binaries aren't zips; nothing to extract.
        except Exception as e:
            self.stats["errors"] += 1
            self.log(f"  ! could not open {full_path} ({e})")

    # -- walk ---------------------------------------------------------------

    def _should_skip_dir(self, path, name):
        if name in SKIP_DIR_NAMES:
            return True
        if name.startswith(OUTPUT_PREFIX):
            return True
        if os.path.realpath(path) == self.output_real:
            return True  # never rescan our own output
        return False

    def scan_root(self, root):
        self.log(f"\nScanning {root} ...")
        last_tick = time.time()
        for dirpath, dirnames, filenames in os.walk(root, topdown=True,
                                                    followlinks=False):
            self.check_stop()
            # Prune skip dirs in place.
            dirnames[:] = [d for d in dirnames
                           if not self._should_skip_dir(os.path.join(dirpath, d), d)]
            for fname in filenames:
                self.check_stop()
                ext = os.path.splitext(fname)[1].lstrip(".").lower()
                full = os.path.join(dirpath, fname)

                if ext in MEDIA_EXTS:
                    self.stats["scanned"] += 1
                    try:
                        tokset, joined = self._path_signals(root, full)
                        score, reasons = self.profile.score(
                            tokset, joined, fuzzy=self.fuzzy)
                        if score >= self.threshold:
                            try:
                                st = os.stat(full)
                                size, mtime = st.st_size, st.st_mtime
                            except Exception:
                                size, mtime = 0, 0
                            category = "image" if ext in IMAGE_EXTS else "video"
                            dest = self._dest_for(root, full)
                            self._copy(full, dest, category, size, mtime, score, reasons)
                            self.stats["media_matched"] += 1
                            if self.stats["media_matched"] <= 5000 and \
                                    self.stats["media_matched"] % 25 == 1:
                                self.log(f"  + {category} ({score:.1f}): "
                                         f"{os.path.relpath(full, root)}")
                    except Stop:
                        raise
                    except Exception as e:
                        self.stats["errors"] += 1
                        self.log(f"  ! error on {full} ({e})")

                elif self.do_pptx and ext in PPTX_EXTS:
                    self._handle_pptx(root, full)

                now = time.time()
                if now - last_tick > 0.4:
                    self.progress()
                    last_tick = now
        self.progress()

    def run(self):
        try:
            for root in self.roots:
                self.check_stop()
                if not os.path.isdir(root):
                    self.log(f"Skipping missing drive: {root}")
                    continue
                self.scan_root(root)
            self._write_manifest()
            self.q.put(("done", dict(self.stats)))
        except Stop:
            self._write_manifest()
            self.q.put(("stopped", dict(self.stats)))
        except Exception:
            self.log("FATAL:\n" + traceback.format_exc())
            self.q.put(("error", dict(self.stats)))

    def _write_manifest(self):
        if not self.manifest_rows:
            return
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            path = os.path.join(self.output_dir, "manifest.csv")
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=[
                    "category", "source", "dest", "size_bytes",
                    "modified", "score", "reasons"])
                w.writeheader()
                for row in self.manifest_rows:
                    w.writerow(row)
            self.log(f"\nManifest written: {path}")
        except Exception as e:
            self.log(f"! could not write manifest ({e})")


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def find_default_projects_file(start_dir):
    """Look for a likely projects .md file next to the app."""
    candidates = []
    try:
        for name in os.listdir(start_dir):
            if name.lower().endswith((".md", ".markdown", ".txt")):
                low = name.lower()
                if low in ("readme.md", "readme.txt", "license.md"):
                    continue
                p = os.path.join(start_dir, name)
                if os.path.isfile(p):
                    score = 0
                    if "project" in low:
                        score += 10
                    if low.endswith(".md"):
                        score += 2
                    try:
                        score += min(os.path.getsize(p) / 1000.0, 50)
                    except Exception:
                        pass
                    candidates.append((score, p))
    except Exception:
        pass
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    return ""


def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext

    app_dir = app_drive_dir()

    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry("860x720")
    root.minsize(720, 600)

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    ttk.Label(main, text=APP_NAME,
              font=("Helvetica", 16, "bold")).pack(anchor="w")
    ttk.Label(main, text="Gathers every image/video that might belong to your "
              "projects into a new folder on this drive. Scanned drives are "
              "only ever read — never changed.",
              wraplength=820, foreground="#555").pack(anchor="w", pady=(0, 8))

    # --- projects file ---
    pf_frame = ttk.LabelFrame(main, text="1. Projects description file (Markdown)",
                              padding=8)
    pf_frame.pack(fill="x", pady=4)
    pf_var = tk.StringVar(value=find_default_projects_file(app_dir))
    pf_row = ttk.Frame(pf_frame)
    pf_row.pack(fill="x")
    pf_entry = ttk.Entry(pf_row, textvariable=pf_var)
    pf_entry.pack(side="left", fill="x", expand=True)

    def browse_pf():
        p = filedialog.askopenfilename(
            title="Choose your projects description file",
            initialdir=app_dir,
            filetypes=[("Text / Markdown", "*.md *.markdown *.txt"),
                       ("All files", "*.*")])
        if p:
            pf_var.set(p)
    ttk.Button(pf_row, text="Browse…", command=browse_pf).pack(side="left", padx=(6, 0))
    pf_info = ttk.Label(pf_frame, text="", foreground="#777")
    pf_info.pack(anchor="w", pady=(4, 0))

    # --- drives ---
    dr_frame = ttk.LabelFrame(main, text="2. Drives to scan", padding=8)
    dr_frame.pack(fill="x", pady=4)
    drives_container = ttk.Frame(dr_frame)
    drives_container.pack(fill="x")
    drive_vars = []

    def refresh_drives():
        for child in drives_container.winfo_children():
            child.destroy()
        drive_vars.clear()
        for label, path in list_drives():
            v = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(drives_container, text=label, variable=v)
            cb.pack(anchor="w")
            drive_vars.append((v, path))
        if not drive_vars:
            ttk.Label(drives_container, text="No drives found.").pack(anchor="w")
    refresh_drives()
    ttk.Button(dr_frame, text="Refresh drive list",
               command=refresh_drives).pack(anchor="w", pady=(6, 0))

    # --- options ---
    opt_frame = ttk.LabelFrame(main, text="3. Options", padding=8)
    opt_frame.pack(fill="x", pady=4)

    row1 = ttk.Frame(opt_frame)
    row1.pack(fill="x")
    ttk.Label(row1, text="Sensitivity:").pack(side="left")
    sens_var = tk.StringVar(value="High recall (catch almost anything)")
    sens_combo = ttk.Combobox(row1, textvariable=sens_var, state="readonly",
                              values=list(SENSITIVITY.keys()), width=34)
    sens_combo.pack(side="left", padx=(6, 16))

    fuzzy_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(row1, text="Fuzzy matching (typos/variants)",
                    variable=fuzzy_var).pack(side="left")

    row2 = ttk.Frame(opt_frame)
    row2.pack(fill="x", pady=(6, 0))
    pptx_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(row2, text="Open PowerPoint files and extract embedded media",
                    variable=pptx_var).pack(side="left", padx=(0, 16))
    scan_only_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(row2, text="Preview only (count matches, don't copy yet)",
                    variable=scan_only_var).pack(side="left")

    # --- output location ---
    out_frame = ttk.Frame(opt_frame)
    out_frame.pack(fill="x", pady=(6, 0))
    ttk.Label(out_frame, text="Output goes to a new folder in:").pack(side="left")
    out_dir_var = tk.StringVar(value=app_dir)
    ttk.Label(out_frame, textvariable=out_dir_var,
              foreground="#0a6").pack(side="left", padx=(6, 6))

    def choose_out():
        p = filedialog.askdirectory(title="Where should the gathered folder go?",
                                    initialdir=out_dir_var.get())
        if p:
            out_dir_var.set(p)
    ttk.Button(out_frame, text="Change…", command=choose_out).pack(side="left")

    # --- controls ---
    ctrl = ttk.Frame(main)
    ctrl.pack(fill="x", pady=(8, 4))
    start_btn = ttk.Button(ctrl, text="Start gathering")
    start_btn.pack(side="left")
    stop_btn = ttk.Button(ctrl, text="Stop", state="disabled")
    stop_btn.pack(side="left", padx=(8, 0))
    open_btn = ttk.Button(ctrl, text="Open output folder", state="disabled")
    open_btn.pack(side="left", padx=(8, 0))

    status_var = tk.StringVar(value="Ready.")
    ttk.Label(main, textvariable=status_var, foreground="#333").pack(anchor="w")

    log_box = scrolledtext.ScrolledText(main, height=14, wrap="word",
                                        font=("Menlo", 10))
    log_box.pack(fill="both", expand=True, pady=(4, 0))
    log_box.configure(state="disabled")

    # --- state ---
    state = {"thread": None, "queue": None, "stop": None,
             "output": None, "profile": None}

    def append_log(text):
        log_box.configure(state="normal")
        log_box.insert("end", text + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")

    def set_running(running):
        start_btn.configure(state="disabled" if running else "normal")
        stop_btn.configure(state="normal" if running else "disabled")
        for v, _ in drive_vars:
            pass  # checkboxes stay visible; harmless if toggled mid-run

    def open_output():
        path = state["output"]
        if not path or not os.path.isdir(path):
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # noqa
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", path])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Could not open folder:\n{e}")
    open_btn.configure(command=open_output)

    def poll_queue():
        q = state["queue"]
        if q is None:
            return
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == "log":
                    append_log(payload)
                elif kind == "progress":
                    s = payload
                    status_var.set(
                        f"Scanned {s['scanned']:,} media · "
                        f"matched {s['media_matched']:,} · "
                        f"PPT decks matched {s['pptx_matched']:,} "
                        f"({s['pptx_media']:,} embedded) · "
                        f"copied {s['copied']:,} ({human_size(s['bytes'])})")
                elif kind in ("done", "stopped", "error"):
                    s = payload
                    verb = {"done": "Finished", "stopped": "Stopped",
                            "error": "Error"}[kind]
                    status_var.set(
                        f"{verb}. Scanned {s['scanned']:,} media, "
                        f"matched {s['media_matched']:,}, "
                        f"{s['pptx_matched']:,} PPT decks "
                        f"({s['pptx_media']:,} embedded media), "
                        f"copied {s['copied']:,} files "
                        f"({human_size(s['bytes'])}), {s['errors']} errors.")
                    append_log(f"\n=== {verb} ===")
                    set_running(False)
                    state["thread"] = None
                    state["queue"] = None
                    if state["output"] and os.path.isdir(state["output"]):
                        open_btn.configure(state="normal")
                    return
        except queue.Empty:
            pass
        root.after(150, poll_queue)

    def start():
        pf = pf_var.get().strip()
        if not pf or not os.path.isfile(pf):
            messagebox.showerror(APP_NAME,
                                 "Please choose your projects description file.")
            return
        roots = [p for v, p in drive_vars if v.get()]
        if not roots:
            messagebox.showerror(APP_NAME, "Please select at least one drive to scan.")
            return
        try:
            with open(pf, "r", encoding="utf-8", errors="ignore") as f:
                md = f.read()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Could not read projects file:\n{e}")
            return

        profile = ProjectProfile().build_from_markdown(md)
        if profile.term_count() == 0:
            messagebox.showerror(
                APP_NAME,
                "No project terms could be extracted from that file.\n"
                "Make sure it contains your project names and descriptions.")
            return
        names = profile.summary()
        pf_info.configure(
            text=f"Extracted {profile.term_count()} project terms. "
                 f"Top names: " + "; ".join(names[:6]))

        # Confirm big scans of the root filesystem.
        risky = [r for r in roots if r in ("/",) or (len(r) <= 3 and r.endswith(":\\"))]
        # (root/system drives are allowed; just proceed.)

        # Unique, context-named output folder.
        host = sanitize(socket.gethostname() or "computer")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        pf_tag = sanitize(os.path.splitext(os.path.basename(pf))[0])[:24]
        folder = f"{OUTPUT_PREFIX}__{host}__{pf_tag}__{stamp}"
        parent = out_dir_var.get().strip() or app_dir
        output = os.path.join(parent, folder)

        copy_enabled = not scan_only_var.get()
        if copy_enabled:
            try:
                os.makedirs(output, exist_ok=True)
            except Exception as e:
                messagebox.showerror(APP_NAME,
                                     f"Cannot create output folder:\n{output}\n\n{e}")
                return
        state["output"] = output

        log_box.configure(state="normal")
        log_box.delete("1.0", "end")
        log_box.configure(state="disabled")
        append_log(f"{APP_NAME}")
        append_log(f"Projects file : {pf}")
        append_log(f"Terms         : {profile.term_count()} "
                   f"({len(profile.phrases)} phrases, {len(profile.tokens)} words)")
        append_log(f"Top names     : " + "; ".join(names))
        append_log(f"Drives        : " + ", ".join(roots))
        append_log(f"Sensitivity   : {sens_var.get()}  "
                   f"(fuzzy={'on' if fuzzy_var.get() else 'off'})")
        append_log(f"Output        : {output}"
                   + ("" if copy_enabled else "   [PREVIEW ONLY - no files copied]"))

        q = queue.Queue()
        stop_event = threading.Event()
        scanner = Scanner(
            profile=profile, roots=roots, output_dir=output,
            sensitivity=SENSITIVITY[sens_var.get()], fuzzy=fuzzy_var.get(),
            do_pptx=pptx_var.get(), copy_enabled=copy_enabled,
            msg_queue=q, stop_event=stop_event)
        thread = threading.Thread(target=scanner.run, daemon=True)
        state.update(queue=q, stop=stop_event, thread=thread)
        set_running(True)
        open_btn.configure(state="disabled")
        status_var.set("Scanning…")
        thread.start()
        root.after(150, poll_queue)

    def stop():
        if state["stop"]:
            state["stop"].set()
            status_var.set("Stopping…")

    start_btn.configure(command=start)
    stop_btn.configure(command=stop)

    def on_close():
        if state["stop"]:
            state["stop"].set()
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()


# ---------------------------------------------------------------------------
# Command-line fallback (headless environments / no Tkinter)
# ---------------------------------------------------------------------------

def run_cli(argv):
    import argparse
    p = argparse.ArgumentParser(
        description="Project Gatherer (command-line mode). "
                    "With no arguments the graphical app is launched.")
    p.add_argument("--projects", required=True, help="Path to projects .md file")
    p.add_argument("--drive", action="append", dest="drives", required=True,
                   help="Drive/folder to scan (repeatable)")
    p.add_argument("--sensitivity", default="high",
                   choices=["high", "balanced", "strict"])
    p.add_argument("--no-fuzzy", action="store_true")
    p.add_argument("--no-pptx", action="store_true")
    p.add_argument("--scan-only", action="store_true",
                   help="Count matches without copying")
    p.add_argument("--output", default=None,
                   help="Parent folder for the gathered output "
                        "(default: next to the app)")
    args = p.parse_args(argv)

    with open(args.projects, "r", encoding="utf-8", errors="ignore") as f:
        profile = ProjectProfile().build_from_markdown(f.read())
    print(f"Extracted {profile.term_count()} project terms "
          f"({len(profile.phrases)} phrases, {len(profile.tokens)} words).")
    print("Top names: " + "; ".join(profile.summary()))

    threshold = {"high": SENSITIVITY["High recall (catch almost anything)"],
                 "balanced": SENSITIVITY["Balanced"],
                 "strict": SENSITIVITY["Strict (high confidence only)"]}[args.sensitivity]

    host = sanitize(socket.gethostname() or "computer")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    pf_tag = sanitize(os.path.splitext(os.path.basename(args.projects))[0])[:24]
    parent = args.output or app_drive_dir()
    output = os.path.join(parent, f"{OUTPUT_PREFIX}__{host}__{pf_tag}__{stamp}")
    if not args.scan_only:
        os.makedirs(output, exist_ok=True)
    print(f"Output: {output}"
          + ("  [PREVIEW ONLY]" if args.scan_only else ""))

    q = queue.Queue()
    stop_event = threading.Event()
    scanner = Scanner(
        profile=profile, roots=args.drives, output_dir=output,
        sensitivity=threshold, fuzzy=not args.no_fuzzy,
        do_pptx=not args.no_pptx, copy_enabled=not args.scan_only,
        msg_queue=q, stop_event=stop_event)

    t = threading.Thread(target=scanner.run, daemon=True)
    t.start()
    while t.is_alive() or not q.empty():
        try:
            kind, payload = q.get(timeout=0.3)
        except queue.Empty:
            continue
        if kind == "log":
            print(payload)
        elif kind == "progress":
            s = payload
            sys.stdout.write(
                f"\r  scanned {s['scanned']:,} matched {s['media_matched']:,} "
                f"copied {s['copied']:,} ({human_size(s['bytes'])})   ")
            sys.stdout.flush()
        elif kind in ("done", "stopped", "error"):
            s = payload
            print(f"\n{kind.upper()}: matched {s['media_matched']:,} media, "
                  f"{s['pptx_matched']:,} PPT decks ({s['pptx_media']:,} embedded), "
                  f"copied {s['copied']:,} ({human_size(s['bytes'])}), "
                  f"{s['errors']} errors.")
    return 0


def main():
    # If CLI args are given, run headless; otherwise launch the GUI.
    if len(sys.argv) > 1:
        return run_cli(sys.argv[1:])
    try:
        run_gui()
        return 0
    except Exception as e:
        print(f"Could not start the graphical interface: {e}")
        print("Run with --help to use the command-line mode instead.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
