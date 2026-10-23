#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Data, Reader};
use eframe::egui::{self, Align2, Color32, Key, Pos2, Rect, ScrollArea, Sense, TextStyle, Vec2};
use std::collections::{BTreeMap, HashSet, VecDeque};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::mpsc::{channel, Receiver};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

const RH: f32 = 20.0;
const HH: f32 = 22.0;
const MAX_HITS: usize = 100_000;
const EXTS: [&str; 6] = ["xlsx", "xlsm", "xlsb", "xls", "ods", "xlam"];

// ---------------------------------------------------------------- data model

struct Sheet {
    name: String,
    row0: usize,
    col0: usize,
    nrows: usize,
    ncols: usize,
    cells: Vec<String>,
    num: Vec<bool>,
    col_x: Vec<f32>,
}

impl Sheet {
    fn cell(&self, r: usize, c: usize) -> &str {
        &self.cells[r * self.ncols + c]
    }
}

struct Book {
    sheets: Vec<Sheet>,
}

enum State {
    Pending,
    Ready(Arc<Book>),
    Failed(String),
}

struct Wb {
    name: String,
    path: PathBuf,
    state: State,
    active_sheet: usize,
}

struct Group {
    label: String,
    wbs: Vec<usize>,
}

fn col_name(mut i: usize) -> String {
    let mut s = Vec::new();
    loop {
        s.push(b'A' + (i % 26) as u8);
        if i < 26 {
            break;
        }
        i = i / 26 - 1;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

fn fmt_float(f: f64) -> String {
    if f.fract() == 0.0 && f.abs() < 1e15 {
        return format!("{}", f as i64);
    }
    let r: f64 = format!("{:.14e}", f).parse().unwrap_or(f);
    format!("{}", r)
}

fn fmt_excel_dt(v: f64, duration: bool) -> String {
    if duration {
        let t = (v * 86400.0).round() as i64;
        return format!("{}:{:02}:{:02}", t / 3600, (t / 60) % 60, t % 60);
    }
    let mut days = v.floor() as i64;
    let mut secs = ((v - v.floor()) * 86400.0).round() as i64;
    if secs >= 86400 {
        days += 1;
        secs = 0;
    }
    let time = format!("{:02}:{:02}:{:02}", secs / 3600, (secs / 60) % 60, secs % 60);
    if days < 1 {
        return time;
    }
    let z = days - 25569 + 719468;
    let era = z.div_euclid(146097);
    let doe = z.rem_euclid(146097);
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = yoe + era * 400 + if m <= 2 { 1 } else { 0 };
    if secs == 0 {
        format!("{:04}-{:02}-{:02}", y, m, d)
    } else {
        format!("{:04}-{:02}-{:02} {}", y, m, d, time)
    }
}

fn load_book(path: &Path) -> Result<Book, String> {
    let path = path.to_path_buf();
    let res = std::panic::catch_unwind(move || {
        let mut wb = open_workbook_auto(&path).map_err(|e| e.to_string())?;
        Ok::<_, String>(wb.worksheets())
    });
    let sheets_raw = match res {
        Ok(r) => r?,
        Err(_) => return Err("file could not be parsed".into()),
    };
    let mut sheets = Vec::with_capacity(sheets_raw.len());
    for (name, range) in sheets_raw {
        let (nrows, ncols) = range.get_size();
        let (row0, col0) = range.start().map(|(r, c)| (r as usize, c as usize)).unwrap_or((0, 0));
        let mut cells = Vec::with_capacity(nrows * ncols);
        let mut num = Vec::with_capacity(nrows * ncols);
        for row in range.rows() {
            for d in row {
                let (s, n) = match d {
                    Data::Empty => (String::new(), false),
                    Data::String(s) => (s.clone(), false),
                    Data::Float(f) => (fmt_float(*f), true),
                    Data::Int(i) => (i.to_string(), true),
                    Data::Bool(b) => ((if *b { "TRUE" } else { "FALSE" }).to_string(), false),
                    Data::DateTime(dt) => (fmt_excel_dt(dt.as_f64(), dt.is_duration()), true),
                    Data::DateTimeIso(s) | Data::DurationIso(s) => (s.clone(), false),
                    Data::Error(e) => (e.to_string(), false),
                };
                cells.push(s);
                num.push(n);
            }
        }
        // column widths from the first rows
        let mut widths = vec![48.0f32; ncols];
        for r in 0..nrows.min(300) {
            for c in 0..ncols {
                let n = cells[r * ncols + c].chars().take(60).count() as f32;
                widths[c] = widths[c].max(n * 7.0 + 14.0);
            }
        }
        let mut col_x = Vec::with_capacity(ncols + 1);
        let mut x = 0.0;
        col_x.push(0.0);
        for w in widths {
            x += w.min(320.0);
            col_x.push(x);
        }
        sheets.push(Sheet { name, row0, col0, nrows, ncols, cells, num, col_x });
    }
    Ok(Book { sheets })
}

// ------------------------------------------------------------------- loader

struct Queue {
    q: Mutex<(VecDeque<usize>, bool)>,
    claimed: Vec<AtomicBool>,
}

struct Loader {
    queue: Arc<Queue>,
    active: Arc<Mutex<Vec<(usize, Instant)>>>,
    rx: Receiver<(usize, Result<Book, String>)>,
}

impl Loader {
    fn start(ctx: &egui::Context, paths: Vec<PathBuf>) -> Loader {
        let n = paths.len();
        let queue = Arc::new(Queue {
            q: Mutex::new(((0..n).collect(), false)),
            claimed: (0..n).map(|_| AtomicBool::new(false)).collect(),
        });
        let (tx, rx) = channel();
        let active: Arc<Mutex<Vec<(usize, Instant)>>> = Arc::default();
        let paths = Arc::new(paths);
        let threads = std::thread::available_parallelism().map(|n| n.get()).unwrap_or(4).clamp(2, 8);
        for _ in 0..threads {
            let (queue, tx, paths, ctx, active) = (queue.clone(), tx.clone(), paths.clone(), ctx.clone(), active.clone());
            std::thread::spawn(move || loop {
                let idx = {
                    let mut g = queue.q.lock().unwrap();
                    if g.1 {
                        return;
                    }
                    match g.0.pop_front() {
                        Some(i) => i,
                        None => return,
                    }
                };
                if queue.claimed[idx].swap(true, Ordering::SeqCst) {
                    continue;
                }
                active.lock().unwrap().push((idx, Instant::now()));
                let r = load_book(&paths[idx]);
                active.lock().unwrap().retain(|a| a.0 != idx);
                if tx.send((idx, r)).is_err() {
                    return;
                }
                ctx.request_repaint();
            });
        }
        Loader { queue, active, rx }
    }

    fn prioritize(&self, idx: usize) {
        if !self.queue.claimed[idx].load(Ordering::SeqCst) {
            self.queue.q.lock().unwrap().0.push_front(idx);
        }
    }
}

impl Drop for Loader {
    fn drop(&mut self) {
        let mut g = self.queue.q.lock().unwrap();
        g.1 = true;
        g.0.clear();
    }
}

// ------------------------------------------------------------------- search

#[derive(Clone, Copy, PartialEq, Eq)]
enum Scope {
    Sheet,
    Workbook,
    Folder,
    All,
}

impl Scope {
    fn label(self) -> &'static str {
        match self {
            Scope::Sheet => "This sheet",
            Scope::Workbook => "This workbook",
            Scope::Folder => "This folder",
            Scope::All => "Everything",
        }
    }
}

#[derive(Clone, Copy)]
struct Hit {
    wb: usize,
    sheet: usize,
    row: usize,
    col: usize,
}

#[derive(Default)]
struct SearchOut {
    gen: u64,
    hits: Vec<Hit>,
    done: bool,
}

struct SearchShared {
    gen: AtomicU64,
    out: Mutex<SearchOut>,
}

fn matches(hay: &str, needle: &str, case: bool) -> bool {
    if case {
        return hay.contains(needle);
    }
    if hay.is_ascii() && needle.is_ascii() {
        let (h, n) = (hay.as_bytes(), needle.as_bytes());
        n.len() <= h.len() && h.windows(n.len()).any(|w| w.eq_ignore_ascii_case(n))
    } else {
        hay.to_lowercase().contains(needle)
    }
}

fn spawn_search(
    shared: Arc<SearchShared>,
    my_gen: u64,
    books: Vec<(usize, Arc<Book>)>,
    only_sheet: Option<usize>,
    query: String,
    case: bool,
    ctx: egui::Context,
) {
    std::thread::spawn(move || {
        let needle = if case { query } else { query.to_lowercase() };
        let mut total = 0usize;
        'outer: for (wi, book) in books {
            for (si, sheet) in book.sheets.iter().enumerate() {
                if only_sheet.is_some_and(|o| o != si) {
                    continue;
                }
                let mut local = Vec::new();
                for r in 0..sheet.nrows {
                    if r % 128 == 0 && shared.gen.load(Ordering::Relaxed) != my_gen {
                        return;
                    }
                    for c in 0..sheet.ncols {
                        let t = sheet.cell(r, c);
                        if !t.is_empty() && matches(t, &needle, case) {
                            local.push(Hit { wb: wi, sheet: si, row: r, col: c });
                            total += 1;
                            if total >= MAX_HITS {
                                break;
                            }
                        }
                    }
                    if total >= MAX_HITS {
                        break;
                    }
                }
                if !local.is_empty() {
                    let mut o = shared.out.lock().unwrap();
                    if o.gen != my_gen {
                        return;
                    }
                    o.hits.extend(local);
                    ctx.request_repaint();
                }
                if total >= MAX_HITS {
                    break 'outer;
                }
            }
        }
        let mut o = shared.out.lock().unwrap();
        if o.gen == my_gen {
            o.done = true;
        }
        ctx.request_repaint();
    });
}

// ---------------------------------------------------------------------- app

struct App {
    root: Option<PathBuf>,
    path_input: String,
    wbs: Vec<Wb>,
    groups: Vec<Group>,
    sel_group: usize,
    sel_wb: Option<usize>,
    loader: Option<Loader>,
    loaded: usize,
    sel: Option<(usize, usize)>,
    scroll_to: Option<Vec2>,
    last_view: Rect,
    last_view_id: (usize, usize),

    query: String,
    case: bool,
    scope: Scope,
    focus_search: bool,
    shared: Arc<SearchShared>,
    hits: Vec<Hit>,
    search_done: bool,
    cur_hit: Option<usize>,
    search_key: Option<(String, bool, Scope, usize, usize, usize, bool)>,
    last_restart: Instant,
    search_loaded: usize,
    view_hits: HashSet<(usize, usize)>,
    vh_key: (usize, bool, usize, usize),
    show_results: bool,
    pal_open: bool,
    pal_query: String,
    pal_sel: usize,
    vertical_wbs: bool,
}

#[derive(Clone, Copy)]
enum Cmd {
    OpenFolder,
    Reload,
    FocusSearch,
    ClearSearch,
    NextHit,
    PrevHit,
    SetScope(Scope),
    ToggleCase,
    ToggleResults,
    CopyCell,
    NextSheet,
    PrevSheet,
    NextWb,
    PrevWb,
    NextFolder,
    PrevFolder,
    GotoWb(usize),
    GotoSheet(usize),
    ToggleTheme,
    ToggleVerticalWbs,
}

fn fuzzy(hay: &str, q: &str) -> Option<usize> {
    let h = hay.to_lowercase();
    let q = q.to_lowercase();
    if q.is_empty() {
        return Some(0);
    }
    if let Some(p) = h.find(&q) {
        return Some(p);
    }
    let mut it = h.chars();
    for qc in q.chars() {
        it.by_ref().find(|&c| c == qc)?;
    }
    Some(1000)
}

fn scan(root: &Path) -> Vec<(PathBuf, String)> {
    let mut out = Vec::new();
    for e in walkdir::WalkDir::new(root).into_iter().filter_map(Result::ok) {
        if !e.file_type().is_file() {
            continue;
        }
        let name = e.file_name().to_string_lossy();
        if name.starts_with("~$") {
            continue;
        }
        let ext = e.path().extension().map(|x| x.to_string_lossy().to_lowercase()).unwrap_or_default();
        if !EXTS.contains(&ext.as_str()) {
            continue;
        }
        let rel = e
            .path()
            .parent()
            .and_then(|p| p.strip_prefix(root).ok())
            .map(|p| p.to_string_lossy().replace('\\', "/"))
            .unwrap_or_default();
        out.push((e.path().to_path_buf(), rel));
    }
    out.sort_by_key(|(p, rel)| (rel.to_lowercase(), p.file_name().map(|n| n.to_string_lossy().to_lowercase())));
    out
}

impl App {
    fn new() -> App {
        App {
            root: None,
            path_input: String::new(),
            wbs: Vec::new(),
            groups: Vec::new(),
            sel_group: 0,
            sel_wb: None,
            loader: None,
            loaded: 0,
            sel: None,
            scroll_to: None,
            last_view: Rect::NOTHING,
            last_view_id: (usize::MAX, 0),
            query: String::new(),
            case: false,
            scope: Scope::All,
            focus_search: false,
            shared: Arc::new(SearchShared { gen: AtomicU64::new(0), out: Mutex::new(SearchOut::default()) }),
            hits: Vec::new(),
            search_done: true,
            cur_hit: None,
            search_key: None,
            last_restart: Instant::now(),
            search_loaded: 0,
            view_hits: HashSet::new(),
            vh_key: (usize::MAX, false, 0, 0),
            show_results: true,
            pal_open: false,
            pal_query: String::new(),
            pal_sel: 0,
            vertical_wbs: false,
        }
    }

    fn open_folder(&mut self, ctx: &egui::Context, path: PathBuf) {
        let files = scan(&path);
        self.loader = None;
        self.wbs.clear();
        self.groups.clear();
        let mut by_dir: BTreeMap<String, Vec<usize>> = BTreeMap::new();
        for (i, (p, rel)) in files.iter().enumerate() {
            self.wbs.push(Wb {
                name: p.file_name().unwrap().to_string_lossy().into_owned(),
                path: p.clone(),
                state: State::Pending,
                active_sheet: 0,
            });
            by_dir.entry(rel.clone()).or_default().push(i);
        }
        self.groups.push(Group { label: "All".into(), wbs: (0..files.len()).collect() });
        for (rel, wbs) in by_dir {
            let label = if rel.is_empty() { "(root)".to_string() } else { rel };
            self.groups.push(Group { label, wbs });
        }
        self.sel_group = 0;
        self.sel_wb = if files.is_empty() { None } else { Some(0) };
        self.sel = None;
        self.loaded = 0;
        self.hits.clear();
        self.search_key = None;
        self.last_view_id = (usize::MAX, 0);
        self.path_input = path.to_string_lossy().into_owned();
        ctx.send_viewport_cmd(egui::ViewportCommand::Title(format!("xlview (read-only) - {}", path.display())));
        self.loader = Some(Loader::start(ctx, files.into_iter().map(|(p, _)| p).collect()));
        self.root = Some(path);
    }

    fn book(&self, wi: usize) -> Option<Arc<Book>> {
        match &self.wbs.get(wi)?.state {
            State::Ready(b) => Some(b.clone()),
            _ => None,
        }
    }

    fn cur_sheet_id(&self) -> Option<(usize, usize)> {
        let wi = self.sel_wb?;
        let b = self.book(wi)?;
        let si = self.wbs[wi].active_sheet.min(b.sheets.len().saturating_sub(1));
        Some((wi, si))
    }

    fn poll_loader(&mut self, ctx: &egui::Context) {
        let Some(l) = &self.loader else { return };
        while let Ok((i, r)) = l.rx.try_recv() {
            self.wbs[i].state = match r {
                Ok(b) => State::Ready(Arc::new(b)),
                Err(e) => State::Failed(e),
            };
            self.loaded += 1;
        }
        if self.loaded >= self.wbs.len() {
            self.loader = None;
        }
        let _ = ctx;
    }

    /// Files currently being parsed, with elapsed seconds (longest-running first).
    fn loading_now(&self) -> Vec<(String, f32)> {
        let Some(l) = &self.loader else { return Vec::new() };
        let mut v: Vec<(String, f32)> =
            l.active.lock().unwrap().iter().map(|(i, t)| (self.wbs[*i].name.clone(), t.elapsed().as_secs_f32())).collect();
        v.sort_by(|a, b| b.1.total_cmp(&a.1));
        v
    }

    fn select_wb(&mut self, wi: usize) {
        if self.sel_wb != Some(wi) {
            self.sel_wb = Some(wi);
            self.sel = None;
        }
    }

    fn select_group(&mut self, gi: usize) {
        self.sel_group = gi;
        let g = &self.groups[gi];
        if !self.sel_wb.is_some_and(|w| g.wbs.contains(&w)) {
            if let Some(&w) = g.wbs.first() {
                self.select_wb(w);
            }
        }
    }

    fn reveal(&mut self, wi: usize, si: usize, r: usize, c: usize, center: bool) {
        let Some(book) = self.book(wi) else { return };
        let Some(sheet) = book.sheets.get(si) else { return };
        let rn_w = rn_width(sheet);
        let (x0, x1) = (rn_w + sheet.col_x[c], rn_w + sheet.col_x[c + 1]);
        let (y0, y1) = (HH + r as f32 * RH, HH + (r + 1) as f32 * RH);
        let v = self.last_view;
        let same = self.last_view_id == (wi, si);
        let visible = same && x0 >= v.min.x + rn_w && x1 <= v.max.x && y0 >= v.min.y + HH && y1 <= v.max.y;
        if visible {
            return;
        }
        let (w, h) = if same { (v.width(), v.height()) } else { (900.0, 500.0) };
        let off = if center || !same {
            Vec2::new((x0 - rn_w - w / 2.0).max(0.0), (y0 - HH - h / 2.0).max(0.0))
        } else {
            let mut o = v.min.to_vec2();
            if x0 < v.min.x + rn_w {
                o.x = x0 - rn_w;
            } else if x1 > v.max.x {
                o.x = x1 - w;
            }
            if y0 < v.min.y + HH {
                o.y = y0 - HH;
            } else if y1 > v.max.y {
                o.y = y1 - h;
            }
            o
        };
        self.scroll_to = Some(off);
    }

    fn goto_hit(&mut self, i: usize) {
        let h = self.hits[i];
        self.cur_hit = Some(i);
        if !self.groups[self.sel_group].wbs.contains(&h.wb) {
            self.sel_group = self.groups.iter().skip(1).position(|g| g.wbs.contains(&h.wb)).map_or(0, |p| p + 1);
        }
        self.sel_wb = Some(h.wb);
        self.wbs[h.wb].active_sheet = h.sheet;
        self.sel = Some((h.row, h.col));
        self.reveal(h.wb, h.sheet, h.row, h.col, true);
    }

    fn step_hit(&mut self, dir: i32) {
        let n = self.hits.len();
        if n == 0 {
            return;
        }
        let next = match self.cur_hit {
            Some(i) => (i as i64 + dir as i64).rem_euclid(n as i64) as usize,
            None => {
                let cur = self.cur_sheet_id();
                let first = cur.and_then(|(w, s)| self.hits.iter().position(|h| h.wb == w && h.sheet == s));
                match (first, dir) {
                    (Some(f), _) => f,
                    (None, d) if d > 0 => 0,
                    _ => n - 1,
                }
            }
        };
        self.goto_hit(next);
    }

    fn maintain_search(&mut self, ctx: &egui::Context) {
        // restart the search when its inputs change
        let cur = self.cur_sheet_id();
        let ready = cur.is_some();
        let (w, s) = cur.unwrap_or((usize::MAX, 0));
        let folder_wbs = if self.scope == Scope::Folder { self.groups.get(self.sel_group).map_or(0, |g| g.wbs.len()) } else { 0 };
        let key = (self.query.clone(), self.case, self.scope, w, s, folder_wbs, ready);
        let loaded = self.loaded;
        let uses_count = matches!(self.scope, Scope::All | Scope::Folder);
        let changed = self.search_key.as_ref() != Some(&key);
        let count_changed = uses_count && self.search_loaded != loaded;
        if !changed && !count_changed {
            return;
        }
        if !changed && self.last_restart.elapsed() < Duration::from_millis(400) {
            ctx.request_repaint_after(Duration::from_millis(400));
            return;
        }
        self.search_key = Some(key);
        self.search_loaded = loaded;
        self.last_restart = Instant::now();
        let gen = self.shared.gen.fetch_add(1, Ordering::SeqCst) + 1;
        {
            let mut o = self.shared.out.lock().unwrap();
            *o = SearchOut { gen, hits: Vec::new(), done: false };
        }
        self.hits.clear();
        self.cur_hit = None;
        if self.query.is_empty() {
            self.search_done = true;
            return;
        }
        self.search_done = false;
        let (wis, only): (Vec<usize>, Option<usize>) = match self.scope {
            Scope::Sheet => (cur.map(|c| c.0).into_iter().collect(), cur.map(|c| c.1)),
            Scope::Workbook => (self.sel_wb.into_iter().collect(), None),
            Scope::Folder => (self.groups.get(self.sel_group).map(|g| g.wbs.clone()).unwrap_or_default(), None),
            Scope::All => ((0..self.wbs.len()).collect(), None),
        };
        let books: Vec<(usize, Arc<Book>)> = wis.into_iter().filter_map(|w| self.book(w).map(|b| (w, b))).collect();
        spawn_search(self.shared.clone(), gen, books, only, self.query.clone(), self.case, ctx.clone());
    }

    fn poll_search(&mut self) {
        let o = self.shared.out.lock().unwrap();
        if o.gen != self.shared.gen.load(Ordering::SeqCst) {
            return;
        }
        if o.hits.len() != self.hits.len() {
            self.hits = o.hits.clone();
        }
        if o.done {
            self.search_done = true;
        }
    }

    fn refresh_view_hits(&mut self) {
        let (w, s) = self.cur_sheet_id().unwrap_or((usize::MAX, 0));
        let key = (self.hits.len(), self.search_done, w, s);
        if key == self.vh_key {
            return;
        }
        self.vh_key = key;
        self.view_hits = self.hits.iter().filter(|h| h.wb == w && h.sheet == s).map(|h| (h.row, h.col)).collect();
    }

    fn palette_items(&self) -> Vec<(String, Cmd)> {
        let mut v: Vec<(String, Cmd)> = vec![
            ("Open folder…".into(), Cmd::OpenFolder),
            ("Reload folder".into(), Cmd::Reload),
            ("Search: focus box".into(), Cmd::FocusSearch),
            ("Search: clear".into(), Cmd::ClearSearch),
            ("Search: next hit".into(), Cmd::NextHit),
            ("Search: previous hit".into(), Cmd::PrevHit),
            ("Search: toggle match case".into(), Cmd::ToggleCase),
            ("Search: toggle results list".into(), Cmd::ToggleResults),
        ];
        for s in [Scope::Sheet, Scope::Workbook, Scope::Folder, Scope::All] {
            v.push((format!("Search scope: {}", s.label()), Cmd::SetScope(s)));
        }
        v.extend([
            ("Copy selected cell".into(), Cmd::CopyCell),
            ("Next sheet".into(), Cmd::NextSheet),
            ("Previous sheet".into(), Cmd::PrevSheet),
            ("Next workbook".into(), Cmd::NextWb),
            ("Previous workbook".into(), Cmd::PrevWb),
            ("Next folder".into(), Cmd::NextFolder),
            ("Previous folder".into(), Cmd::PrevFolder),
            ("Toggle light/dark theme".into(), Cmd::ToggleTheme),
            ("Toggle vertical workbook list".into(), Cmd::ToggleVerticalWbs),
        ]);
        if let Some(wi) = self.sel_wb {
            if let State::Ready(b) = &self.wbs[wi].state {
                for (si, s) in b.sheets.iter().enumerate() {
                    v.push((format!("Go to sheet: {}", s.name), Cmd::GotoSheet(si)));
                }
            }
        }
        for (wi, w) in self.wbs.iter().enumerate() {
            v.push((format!("Go to workbook: {}", w.name), Cmd::GotoWb(wi)));
        }
        v
    }

    fn step_in(list: &[usize], cur: Option<usize>, dir: i32) -> Option<usize> {
        if list.is_empty() {
            return None;
        }
        let n = list.len() as i64;
        let i = cur.and_then(|c| list.iter().position(|&x| x == c)).map_or(if dir > 0 { 0 } else { n - 1 }, |p| (p as i64 + dir as i64).rem_euclid(n));
        Some(list[i as usize])
    }

    fn run_cmd(&mut self, ctx: &egui::Context, cmd: Cmd) {
        match cmd {
            Cmd::OpenFolder => self.pick_folder(ctx),
            Cmd::Reload => {
                if let Some(r) = self.root.clone() {
                    self.open_folder(ctx, r);
                }
            }
            Cmd::FocusSearch => self.focus_search = true,
            Cmd::ClearSearch => self.query.clear(),
            Cmd::NextHit => self.step_hit(1),
            Cmd::PrevHit => self.step_hit(-1),
            Cmd::SetScope(s) => self.scope = s,
            Cmd::ToggleCase => self.case = !self.case,
            Cmd::ToggleResults => self.show_results = !self.show_results,
            Cmd::CopyCell => {
                if let (Some((wi, si)), Some((r, c))) = (self.cur_sheet_id(), self.sel) {
                    if let Some(b) = self.book(wi) {
                        ctx.copy_text(b.sheets[si].cell(r, c).to_string());
                    }
                }
            }
            Cmd::NextSheet | Cmd::PrevSheet => {
                let d = if matches!(cmd, Cmd::NextSheet) { 1 } else { -1 };
                if let Some((wi, si)) = self.cur_sheet_id() {
                    let n = self.book(wi).map_or(1, |b| b.sheets.len()) as i64;
                    self.wbs[wi].active_sheet = (si as i64 + d).rem_euclid(n) as usize;
                    self.sel = None;
                }
            }
            Cmd::NextWb | Cmd::PrevWb => {
                let d = if matches!(cmd, Cmd::NextWb) { 1 } else { -1 };
                if let Some(g) = self.groups.get(self.sel_group) {
                    if let Some(w) = Self::step_in(&g.wbs, self.sel_wb, d) {
                        self.select_wb(w);
                    }
                }
            }
            Cmd::NextFolder | Cmd::PrevFolder => {
                let d = if matches!(cmd, Cmd::NextFolder) { 1 } else { -1 };
                let n = self.groups.len() as i64;
                if n > 0 {
                    self.select_group((self.sel_group as i64 + d).rem_euclid(n) as usize);
                }
            }
            Cmd::GotoWb(wi) => {
                if !self.groups[self.sel_group].wbs.contains(&wi) {
                    self.sel_group = self.groups.iter().skip(1).position(|g| g.wbs.contains(&wi)).map_or(0, |p| p + 1);
                }
                self.select_wb(wi);
            }
            Cmd::GotoSheet(si) => {
                if let Some(wi) = self.sel_wb {
                    self.wbs[wi].active_sheet = si;
                    self.sel = None;
                }
            }
            Cmd::ToggleVerticalWbs => self.vertical_wbs = !self.vertical_wbs,
            Cmd::ToggleTheme => {
                let dark = ctx.style().visuals.dark_mode;
                ctx.set_visuals(if dark { egui::Visuals::light() } else { egui::Visuals::dark() });
            }
        }
    }

    fn ui_palette(&mut self, ctx: &egui::Context) {
        if !self.pal_open {
            return;
        }
        let mut items: Vec<(usize, String, Cmd)> =
            self.palette_items().into_iter().filter_map(|(l, c)| fuzzy(&l, &self.pal_query).map(|s| (s, l, c))).collect();
        items.sort_by_key(|i| i.0);
        let n = items.len();
        let (down, up, enter, esc) = ctx.input(|i| (i.key_pressed(Key::ArrowDown), i.key_pressed(Key::ArrowUp), i.key_pressed(Key::Enter), i.key_pressed(Key::Escape)));
        if down && n > 0 {
            self.pal_sel = (self.pal_sel + 1) % n;
        }
        if up && n > 0 {
            self.pal_sel = (self.pal_sel + n - 1) % n;
        }
        self.pal_sel = self.pal_sel.min(n.saturating_sub(1));
        let mut run = None;
        egui::Window::new("palette")
            .title_bar(false)
            .resizable(false)
            .anchor(Align2::CENTER_TOP, [0.0, 70.0])
            .fixed_size([560.0, 0.0])
            .show(ctx, |ui| {
                let r = ui.add(egui::TextEdit::singleline(&mut self.pal_query).desired_width(f32::INFINITY).hint_text("Type a command…"));
                r.request_focus();
                if r.changed() {
                    self.pal_sel = 0;
                }
                ScrollArea::vertical().max_height(320.0).auto_shrink([false, true]).show(ui, |ui| {
                    for (i, (_, label, cmd)) in items.iter().enumerate() {
                        let resp = ui.selectable_label(i == self.pal_sel, label);
                        if (down || up) && i == self.pal_sel {
                            resp.scroll_to_me(None);
                        }
                        if resp.clicked() {
                            run = Some(*cmd);
                        }
                    }
                });
            });
        if enter && n > 0 {
            run = Some(items[self.pal_sel].2);
        }
        if esc || run.is_some() {
            self.pal_open = false;
        }
        if let Some(c) = run {
            self.run_cmd(ctx, c);
        }
    }

    fn pick_folder(&mut self, ctx: &egui::Context) {
        if let Some(p) = rfd::FileDialog::new().pick_folder() {
            self.open_folder(ctx, p);
        }
    }

    fn shortcuts(&mut self, ctx: &egui::Context) {
        let (ctrl_f, ctrl_o, f3, shift) =
            ctx.input(|i| (i.modifiers.command && i.key_pressed(Key::F), i.modifiers.command && i.key_pressed(Key::O), i.key_pressed(Key::F3), i.modifiers.shift));
        if ctx.input(|i| i.modifiers.command && i.key_pressed(Key::P)) {
            self.pal_open = !self.pal_open;
            self.pal_query.clear();
            self.pal_sel = 0;
        }
        if ctrl_f {
            self.focus_search = true;
        }
        if ctrl_o {
            self.pick_folder(ctx);
        }
        if f3 {
            self.step_hit(if shift { -1 } else { 1 });
        }
        if ctx.wants_keyboard_input() {
            return;
        }
        let Some((wi, si)) = self.cur_sheet_id() else { return };
        let Some(book) = self.book(wi) else { return };
        let sheet = &book.sheets[si];
        if sheet.nrows == 0 || sheet.ncols == 0 {
            return;
        }
        let copy = ctx.input(|i| i.events.iter().any(|e| matches!(e, egui::Event::Copy)));
        if let (true, Some((r, c))) = (copy, self.sel) {
            ctx.copy_text(sheet.cell(r, c).to_string());
        }
        let (dr, dc) = ctx.input(|i| {
            let mut d = (0i64, 0i64);
            if i.key_pressed(Key::ArrowDown) { d.0 += 1 }
            if i.key_pressed(Key::ArrowUp) { d.0 -= 1 }
            if i.key_pressed(Key::ArrowRight) { d.1 += 1 }
            if i.key_pressed(Key::ArrowLeft) { d.1 -= 1 }
            if i.key_pressed(Key::PageDown) { d.0 += 30 }
            if i.key_pressed(Key::PageUp) { d.0 -= 30 }
            d
        });
        if dr != 0 || dc != 0 {
            let (r, c) = self.sel.unwrap_or((0, 0));
            let nr = (r as i64 + dr).clamp(0, sheet.nrows as i64 - 1) as usize;
            let nc = (c as i64 + dc).clamp(0, sheet.ncols as i64 - 1) as usize;
            self.sel = Some((nr, nc));
            self.reveal(wi, si, nr, nc, false);
        }
    }

    fn handle_drop(&mut self, ctx: &egui::Context) {
        let dropped = ctx.input(|i| i.raw.dropped_files.first().and_then(|f| f.path.clone()));
        if let Some(p) = dropped {
            let dir = if p.is_dir() { Some(p) } else { p.parent().map(|d| d.to_path_buf()) };
            if let Some(d) = dir {
                self.open_folder(ctx, d);
            }
        }
    }

    // ---------------------------------------------------------------- panels

    fn ui_top(&mut self, ctx: &egui::Context) {
        let mut want_pick = false;
        let mut want_open: Option<PathBuf> = None;
        let mut step = 0;
        let mut pick_group = None;
        let mut pick_wb = None;
        let mut pick_sheet = None;
        egui::TopBottomPanel::top("top").show(ctx, |ui| {
            ui.horizontal(|ui| {
                if ui.button("Open folder…").clicked() {
                    want_pick = true;
                }
                let r = ui.add(egui::TextEdit::singleline(&mut self.path_input).desired_width(280.0).hint_text("folder path, then Enter"));
                if r.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter)) {
                    let p = PathBuf::from(self.path_input.trim().trim_matches('"'));
                    if p.is_dir() {
                        want_open = Some(p);
                    }
                }
                ui.separator();
                ui.label("Search:");
                let sr = ui.add(egui::TextEdit::singleline(&mut self.query).id(egui::Id::new("search_box")).desired_width(240.0).hint_text("Ctrl+F"));
                if self.focus_search {
                    sr.request_focus();
                    self.focus_search = false;
                }
                if sr.lost_focus() && ui.input(|i| i.key_pressed(Key::Enter)) {
                    step = if ui.input(|i| i.modifiers.shift) { -1 } else { 1 };
                    sr.request_focus();
                }
                ui.toggle_value(&mut self.case, "Aa").on_hover_text("Match case");
                egui::ComboBox::from_id_salt("scope").selected_text(self.scope.label()).show_ui(ui, |ui| {
                    for s in [Scope::Sheet, Scope::Workbook, Scope::Folder, Scope::All] {
                        ui.selectable_value(&mut self.scope, s, s.label());
                    }
                });
                if ui.button("Prev").clicked() {
                    step = -1;
                }
                if ui.button("Next").clicked() {
                    step = 1;
                }
                if !self.query.is_empty() {
                    let n = self.hits.len();
                    let cap = if n >= MAX_HITS { "+" } else { "" };
                    let cur = self.cur_hit.map_or(0, |i| i + 1);
                    ui.label(format!("{}/{}{} {}", cur, n, cap, if self.search_done { "" } else { "…" }));
                    ui.toggle_value(&mut self.show_results, "Results list");
                }
                ui.separator();
                ui.toggle_value(&mut self.vertical_wbs, "Vertical workbooks").on_hover_text("Show workbooks as a vertical list on the left");
            });
            ui.separator();

            // row 1: subfolders
            tab_row(ui, "Folders", "r1", 46.0, |ui| {
                for (i, g) in self.groups.iter().enumerate() {
                    if ui.selectable_label(i == self.sel_group, format!("{} ({})", g.label, g.wbs.len())).clicked() {
                        pick_group = Some(i);
                    }
                }
            });
            // row 2: workbooks of the selected folder (horizontal mode only)
            if !self.vertical_wbs {
                tab_row(ui, "Workbooks", "r2", 66.0, |ui| {
                    if let Some(g) = self.groups.get(self.sel_group) {
                        for &wi in &g.wbs {
                            if wb_label(ui, &self.wbs[wi], self.sel_wb == Some(wi)) {
                                pick_wb = Some(wi);
                            }
                        }
                    }
                });
            }
            // row 3: sheets of the selected workbook
            tab_row(ui, "Sheets", "r3", 46.0, |ui| {
                if let Some(wi) = self.sel_wb {
                    if let State::Ready(b) = &self.wbs[wi].state {
                        let active = self.wbs[wi].active_sheet;
                        for (si, s) in b.sheets.iter().enumerate() {
                            if ui.selectable_label(si == active, &s.name).clicked() {
                                pick_sheet = Some(si);
                            }
                        }
                    }
                }
            });
        });
        if want_pick {
            self.pick_folder(ctx);
        }
        if let Some(p) = want_open {
            self.open_folder(ctx, p);
        }
        if let Some(g) = pick_group {
            self.select_group(g);
        }
        if let Some(w) = pick_wb {
            self.select_wb(w);
        }
        if let (Some(s), Some(w)) = (pick_sheet, self.sel_wb) {
            self.wbs[w].active_sheet = s;
            self.sel = None;
        }
        if step != 0 {
            self.step_hit(step);
        }
    }

    fn ui_wb_side(&mut self, ctx: &egui::Context) {
        if !self.vertical_wbs {
            return;
        }
        let mut pick = None;
        egui::SidePanel::left("wb_list").resizable(true).default_width(240.0).show(ctx, |ui| {
            ui.label(egui::RichText::new("Workbooks").weak());
            ui.separator();
            ScrollArea::both().auto_shrink([false, false]).show(ui, |ui| {
                if let Some(g) = self.groups.get(self.sel_group) {
                    for &wi in &g.wbs {
                        if wb_label(ui, &self.wbs[wi], self.sel_wb == Some(wi)) {
                            pick = Some(wi);
                        }
                    }
                }
            });
        });
        if let Some(w) = pick {
            self.select_wb(w);
        }
    }

    fn ui_status(&mut self, ctx: &egui::Context) {
        egui::TopBottomPanel::bottom("status").show(ctx, |ui| {
            ui.horizontal(|ui| {
                let total = self.wbs.len();
                ui.label(format!("{} workbooks, {}/{} loaded", total, self.loaded.min(total), total));
                let busy = self.loading_now();
                if !busy.is_empty() {
                    ui.add(egui::ProgressBar::new(self.loaded as f32 / total.max(1) as f32).desired_width(100.0));
                    let mut t: Vec<String> = busy.iter().take(3).map(|(n, s)| format!("{} ({:.1}s)", n, s)).collect();
                    if busy.len() > 3 {
                        t.push(format!("+{} more", busy.len() - 3));
                    }
                    ui.add(egui::Label::new(format!("Parsing: {}", t.join(", "))).truncate());
                }
                if let Some((wi, si)) = self.cur_sheet_id() {
                    let b = self.book(wi).unwrap();
                    let s = &b.sheets[si];
                    ui.separator();
                    ui.label(format!("{} rows × {} cols", s.nrows, s.ncols));
                    if let Some((r, c)) = self.sel {
                        if r < s.nrows && c < s.ncols {
                            ui.separator();
                            let t = s.cell(r, c);
                            let t: String = t.chars().take(300).map(|ch| if ch == '\n' { ' ' } else { ch }).collect();
                            ui.add(egui::Label::new(format!("{}{}: {}", col_name(s.col0 + c), s.row0 + r + 1, t)).truncate());
                        }
                    }
                }
            });
        });
    }

    fn ui_results(&mut self, ctx: &egui::Context) {
        if self.query.is_empty() || !self.show_results || self.hits.is_empty() {
            return;
        }
        let mut goto = None;
        egui::TopBottomPanel::bottom("results").resizable(true).default_height(160.0).show(ctx, |ui| {
            let rh = ui.text_style_height(&TextStyle::Body) + ui.spacing().item_spacing.y;
            let (hits, wbs, cur) = (&self.hits, &self.wbs, self.cur_hit);
            ScrollArea::vertical().auto_shrink([false, false]).show_rows(ui, rh, hits.len(), |ui, range| {
                for i in range {
                    let h = hits[i];
                    let wb = &wbs[h.wb];
                    let mut line = String::new();
                    if let State::Ready(b) = &wb.state {
                        let s = &b.sheets[h.sheet];
                        let t: String = s.cell(h.row, h.col).chars().take(120).map(|c| if c == '\n' { ' ' } else { c }).collect();
                        line = format!("{} › {} › {}{}   {}", wb.name, s.name, col_name(s.col0 + h.col), s.row0 + h.row + 1, t);
                    }
                    if ui.selectable_label(cur == Some(i), line).clicked() {
                        goto = Some(i);
                    }
                }
            });
        });
        if let Some(i) = goto {
            self.goto_hit(i);
        }
    }

    fn ui_central(&mut self, ctx: &egui::Context) {
        let busy = self.loading_now();
        egui::CentralPanel::default().show(ctx, |ui| {
            let Some(wi) = self.sel_wb else {
                ui.centered_and_justified(|ui| {
                    ui.label(if self.root.is_some() { "No Excel files found in this folder." } else { "Drop a folder here, press Ctrl+O, or use “Open folder…”" });
                });
                return;
            };
            if let Some(l) = &self.loader {
                l.prioritize(wi);
                ctx.request_repaint_after(Duration::from_millis(150));
            }
            match &self.wbs[wi].state {
                State::Pending => {
                    let name = self.wbs[wi].name.clone();
                    let mine = busy.iter().find(|b| b.0 == name).map(|b| b.1);
                    ui.vertical_centered(|ui| {
                        ui.add_space(ui.available_height() / 3.0);
                        ui.spinner();
                        ui.add_space(8.0);
                        ui.label(egui::RichText::new(&name).strong());
                        match mine {
                            Some(s) => ui.label(format!("Parsing… {:.1}s", s)),
                            None => ui.label("Queued — will start as soon as a worker is free"),
                        };
                        ui.label(format!("{}/{} workbooks loaded", self.loaded, self.wbs.len()));
                        if !busy.is_empty() {
                            ui.add_space(6.0);
                            ui.label(egui::RichText::new("Currently parsing:").weak());
                            for (n, s) in busy.iter().take(8) {
                                ui.label(format!("{} ({:.1}s)", n, s));
                            }
                        }
                    });
                }
                State::Failed(e) => {
                    ui.label(format!("Could not open {}: {}", self.wbs[wi].name, e));
                }
                State::Ready(book) => {
                    let book = book.clone();
                    if book.sheets.is_empty() {
                        ui.label("Workbook has no sheets.");
                        return;
                    }
                    let si = self.wbs[wi].active_sheet.min(book.sheets.len() - 1);
                    self.draw_grid(ui, wi, si, &book.sheets[si]);
                }
            }
        });
    }

    fn draw_grid(&mut self, ui: &mut egui::Ui, wi: usize, si: usize, sheet: &Sheet) {
        if sheet.nrows == 0 || sheet.ncols == 0 {
            ui.label("Empty sheet.");
            return;
        }
        let rn_w = rn_width(sheet);
        let total_w = *sheet.col_x.last().unwrap();
        let content = Vec2::new(rn_w + total_w, HH + sheet.nrows as f32 * RH);
        let mut sa = ScrollArea::both().id_salt(("grid", wi, si)).auto_shrink([false, false]);
        if let Some(o) = self.scroll_to.take() {
            sa = sa.scroll_offset(o);
        }
        self.refresh_view_hits();
        let vis = ui.visuals().clone();
        let font = TextStyle::Body.resolve(ui.style());
        let grid_c = vis.widgets.noninteractive.bg_stroke.color;
        let text_c = vis.text_color();
        let head_bg = vis.widgets.inactive.bg_fill;
        let sel_bg = vis.selection.bg_fill;
        let hit_bg = Color32::from_rgba_unmultiplied(255, 220, 0, 90);
        let cur_hit_bg = Color32::from_rgba_unmultiplied(255, 140, 0, 170);
        let cur_hit = self.cur_hit.map(|i| self.hits[i]).filter(|h| h.wb == wi && h.sheet == si).map(|h| (h.row, h.col));
        let mut sel = self.sel;

        let out = sa.show_viewport(ui, |ui, viewport| {
            let (rect, resp) = ui.allocate_exact_size(content, Sense::click());
            let vis_rect = Rect::from_min_size(rect.min + viewport.min.to_vec2(), viewport.size());
            let painter = ui.painter_at(vis_rect);
            let body = Rect::from_min_max(vis_rect.min + Vec2::new(rn_w, HH), vis_rect.max);
            let bp = painter.with_clip_rect(body);

            let r0 = ((viewport.min.y / RH).floor().max(0.0) as usize).min(sheet.nrows);
            let r1 = (((viewport.max.y - HH) / RH).ceil().max(0.0) as usize + 1).min(sheet.nrows);
            let c0 = sheet.col_x.partition_point(|&x| x <= viewport.min.x).saturating_sub(1).min(sheet.ncols);
            let c1 = sheet.col_x.partition_point(|&x| x < viewport.max.x - rn_w).min(sheet.ncols);

            // click -> select
            if resp.clicked() || resp.double_clicked() {
                if let Some(p) = resp.interact_pointer_pos() {
                    if p.x > vis_rect.min.x + rn_w && p.y > vis_rect.min.y + HH {
                        let rel = p - rect.min;
                        let r = ((rel.y - HH) / RH).floor() as usize;
                        let cx = rel.x - rn_w;
                        let c = sheet.col_x.partition_point(|&x| x <= cx).saturating_sub(1);
                        if r < sheet.nrows && c < sheet.ncols {
                            sel = Some((r, c));
                        }
                    }
                }
            }

            let origin = rect.min;
            for r in r0..r1 {
                let y = origin.y + HH + r as f32 * RH;
                if r % 2 == 1 {
                    bp.rect_filled(Rect::from_min_max(Pos2::new(body.min.x, y), Pos2::new(body.max.x, y + RH)), 0.0, vis.faint_bg_color);
                }
                bp.line_segment([Pos2::new(body.min.x, y + RH), Pos2::new(body.max.x, y + RH)], (1.0, grid_c));
                for c in c0..c1 {
                    let x0 = origin.x + rn_w + sheet.col_x[c];
                    let x1 = origin.x + rn_w + sheet.col_x[c + 1];
                    let cell = Rect::from_min_max(Pos2::new(x0, y), Pos2::new(x1, y + RH));
                    if cur_hit == Some((r, c)) {
                        bp.rect_filled(cell, 0.0, cur_hit_bg);
                    } else if self.view_hits.contains(&(r, c)) {
                        bp.rect_filled(cell, 0.0, hit_bg);
                    }
                    if sel == Some((r, c)) {
                        bp.rect_filled(cell, 0.0, sel_bg);
                    }
                    let t = sheet.cell(r, c);
                    if !t.is_empty() {
                        let cp = bp.with_clip_rect(cell.shrink2(Vec2::new(0.0, 0.0)).intersect(body));
                        let shown: String = t.chars().take(200).map(|ch| if ch == '\n' || ch == '\r' { ' ' } else { ch }).collect();
                        if sheet.num[r * sheet.ncols + c] {
                            cp.text(Pos2::new(x1 - 4.0, y + RH / 2.0), Align2::RIGHT_CENTER, shown, font.clone(), text_c);
                        } else {
                            cp.text(Pos2::new(x0 + 4.0, y + RH / 2.0), Align2::LEFT_CENTER, shown, font.clone(), text_c);
                        }
                    }
                    bp.line_segment([Pos2::new(x1, y), Pos2::new(x1, y + RH)], (1.0, grid_c));
                }
            }

            // sticky column header
            let hdr = Rect::from_min_max(vis_rect.min, Pos2::new(vis_rect.max.x, vis_rect.min.y + HH));
            painter.rect_filled(hdr, 0.0, head_bg);
            let hp = painter.with_clip_rect(Rect::from_min_max(Pos2::new(body.min.x, hdr.min.y), hdr.max));
            for c in c0..c1 {
                let x0 = origin.x + rn_w + sheet.col_x[c];
                let x1 = origin.x + rn_w + sheet.col_x[c + 1];
                if sel.is_some_and(|s| s.1 == c) {
                    hp.rect_filled(Rect::from_min_max(Pos2::new(x0, hdr.min.y), Pos2::new(x1, hdr.max.y)), 0.0, sel_bg);
                }
                hp.text(Pos2::new((x0 + x1) / 2.0, hdr.center().y), Align2::CENTER_CENTER, col_name(sheet.col0 + c), font.clone(), text_c);
                hp.line_segment([Pos2::new(x1, hdr.min.y), Pos2::new(x1, hdr.max.y)], (1.0, grid_c));
            }
            // sticky row numbers
            let rn = Rect::from_min_max(Pos2::new(vis_rect.min.x, vis_rect.min.y + HH), Pos2::new(vis_rect.min.x + rn_w, vis_rect.max.y));
            painter.rect_filled(rn, 0.0, head_bg);
            let rp = painter.with_clip_rect(rn);
            for r in r0..r1 {
                let y = origin.y + HH + r as f32 * RH;
                if sel.is_some_and(|s| s.0 == r) {
                    rp.rect_filled(Rect::from_min_max(Pos2::new(rn.min.x, y), Pos2::new(rn.max.x, y + RH)), 0.0, sel_bg);
                }
                rp.text(Pos2::new(rn.max.x - 6.0, y + RH / 2.0), Align2::RIGHT_CENTER, (sheet.row0 + r + 1).to_string(), font.clone(), text_c);
                rp.line_segment([Pos2::new(rn.min.x, y + RH), Pos2::new(rn.max.x, y + RH)], (1.0, grid_c));
            }
            painter.line_segment([Pos2::new(rn.max.x, hdr.min.y), Pos2::new(rn.max.x, vis_rect.max.y)], (1.0, grid_c));
            painter.line_segment([Pos2::new(hdr.min.x, hdr.max.y), Pos2::new(hdr.max.x, hdr.max.y)], (1.0, grid_c));
            viewport
        });
        self.last_view = out.inner;
        self.last_view_id = (wi, si);
        self.sel = sel;
    }
}

fn rn_width(sheet: &Sheet) -> f32 {
    (sheet.row0 + sheet.nrows).to_string().len() as f32 * 8.0 + 18.0
}

fn wb_label(ui: &mut egui::Ui, wb: &Wb, selected: bool) -> bool {
    let mark = match wb.state {
        State::Pending => " …",
        State::Failed(_) => " !",
        State::Ready(_) => "",
    };
    let r = ui.selectable_label(selected, format!("{}{}", wb.name, mark));
    if let State::Failed(e) = &wb.state {
        r.clone().on_hover_text(e);
    } else {
        r.clone().on_hover_text(wb.path.to_string_lossy());
    }
    r.clicked()
}

fn tab_row(ui: &mut egui::Ui, title: &str, id: &str, max_h: f32, add: impl FnOnce(&mut egui::Ui)) {
    ui.horizontal_top(|ui| {
        ui.add_sized([70.0, 18.0], egui::Label::new(egui::RichText::new(title).weak()));
        ScrollArea::vertical().id_salt(id).max_height(max_h).auto_shrink([false, true]).show(ui, |ui| {
            ui.horizontal_wrapped(add);
        });
    });
}

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.poll_loader(ctx);
        self.handle_drop(ctx);
        self.shortcuts(ctx);
        self.maintain_search(ctx);
        self.poll_search();
        self.ui_top(ctx);
        self.ui_palette(ctx);
        self.ui_status(ctx);
        self.ui_results(ctx);
        self.ui_wb_side(ctx);
        self.ui_central(ctx);
    }
}

fn setup_fonts(ctx: &egui::Context) {
    let mut fonts = egui::FontDefinitions::default();
    for (key, file) in [("yahei", "C:\\Windows\\Fonts\\msyh.ttc"), ("segoe", "C:\\Windows\\Fonts\\segoeui.ttf")] {
        if let Ok(bytes) = std::fs::read(file) {
            fonts.font_data.insert(key.into(), Arc::new(egui::FontData::from_owned(bytes)));
            for fam in [egui::FontFamily::Proportional, egui::FontFamily::Monospace] {
                fonts.families.get_mut(&fam).unwrap().push(key.into());
            }
        }
    }
    ctx.set_fonts(fonts);
}

fn main() -> eframe::Result {
    let arg = std::env::args_os().nth(1).map(PathBuf::from).filter(|p| p.is_dir());
    let opts = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([1400.0, 850.0]).with_title("xlview (read-only)"),
        ..Default::default()
    };
    eframe::run_native(
        "xlview",
        opts,
        Box::new(move |cc| {
            setup_fonts(&cc.egui_ctx);
            let mut app = App::new();
            if let Some(p) = arg {
                app.open_folder(&cc.egui_ctx, p);
            }
            Ok(Box::new(app))
        }),
    )
}
