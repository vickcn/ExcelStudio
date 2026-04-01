#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lightweight local utilities for tmp2/dataProcess.

This module intentionally provides a small LOGger-compatible surface so
dataProcess scripts can run without depending on ../package/LOGger.py.
"""

from __future__ import annotations

import json
import os
import sys
import time
import traceback
from copy import deepcopy as dcp
from datetime import datetime as dt
from pathlib import Path
from typing import Any, Callable, Optional
import chardet as chd

try:
    import numpy as np  # type: ignore
except Exception:  # pragma: no cover
    np = None  # type: ignore


HEADER = '\033[95m'
OKBLUE = '\033[94m'
OKCYAN = '\033[96m'
OKGREEN = '\033[92m'
WARNING = '\033[93m'
FAIL = '\033[91m' #'\x1b[31m' Fore.RED
ENDC = '\033[0m'
BOLD = '\033[1m'
UNDERLINE = '\033[4m'


WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
OUTPUTS_ROOT = WORKSPACE_ROOT / "outputs"


def _ensure_outputs_root() -> None:
    OUTPUTS_ROOT.mkdir(parents=True, exist_ok=True)


def _to_output_path(path_like: Any) -> Path:
    """
    Normalize a path to tmp2/outputs for relative paths.

    Rules:
    - absolute path: keep as-is
    - relative `outputs/...`: map to tmp2/outputs/...
    - relative `log/...`: map to tmp2/outputs/log/...
    - other relative paths: map to tmp2/outputs/<path>
    """
    raw = str(path_like).strip()
    if "%t" in raw:
        raw = raw.replace("%t", dt.now().strftime("%Y%m%d"))

    p = Path(raw)
    if p.is_absolute():
        return p

    raw_norm = raw.replace("\\", "/")
    if raw_norm.startswith("outputs/"):
        return WORKSPACE_ROOT / p
    if raw_norm.startswith("log/"):
        return OUTPUTS_ROOT / p
    return OUTPUTS_ROOT / p


def for_file_process(stamps):
    if(isinstance(stamps, list)):
        stamps = list(map(lambda s:s.replace('<','〈').replace('>','〉'), stamps))
    elif(isinstance(stamps, dict)):
        stamps = {k.replace('<','〈').replace('>','〉'):v.replace('<','〈').replace('>','〉') for (k,v) in stamps.items()}
    elif(isinstance(stamps, str)):
        stamps = str(stamps).replace_all('<','〈').replace_all('>','〉')
    return stamps
#TODO:type_string
def type_string(obj, index=-1, sep='.'):
    a_type = str(type(obj))
    a_type1 = a_type.brackets("<class '", "'>")
    return str(a_type1.split(sep)[index] if(index!=None) else a_type1)


def parse(x, digit=2, stg_max_length=200, be_instinct=False, default_value=None, 
          conjunction=None, max_item_count=10, **kwags):
    try:
        if(isiterable(x)):
            if(hasattr(x, 'shape')):
                return stamp_process('',[type_string(x), str(x.shape)])
            elif(np.array(x).shape[0]<=max_item_count):
                return '%s|%s'%(type_string(x), '`%s`'%stamp_process('',list(map(str, x)),'','','',', '))
            else:
                return stamp_process('',[type_string(x), len(x)])
        elif(isinstance(x, dict)):
            return stamp_process('',[type_string(x), len(x)])
        elif(isnonnumber(x)):
            return str(x)[:stg_max_length]
        else:
            if(be_instinct):
                return int(x) if(float(x)//1==float(x)) else float(('%%.%df'%digit)%float(x))
            else:
                return '%d'%float(x) if(float(x)//1==float(x)) else ('%%.%df'%digit)%float(x)
    except:
        print('parse fail...:%s'%str(x)[:stg_max_length])
        return (x if(be_instinct) else str(x)) if(default_value=='self') else default_value


def isinstance_not_empty(value, _type):
    if(isinstance(_type, tuple)):
        for x in _type:
            if(isinstance_not_empty(value, x)):
                return True
        return False
    return (isinstance(value, _type) if(isinstance(_type, type)) else bool(_type(value))) and (not(not value))

def stamp_process(stg='', stamps=[], stamp_sep=':', stamp_left='[', stamp_right=']', 
                  adjoint_sep='', outer_stamp_left='', outer_stamp_right='', location=1, 
                  exceptions = [''], annih_when_stamps_empty=True, for_file=False, max_len=200, digit=2, **kwags):
    stamp = ''
    if(isinstance(stamps, dict)):
        stamp = adjoint_sep.join(list(map(lambda t:('%s%s%s%s%s'%(stamp_left,
                parse(t[0], stg_max_length=200, digit=digit), stamp_sep, parse(
                    t[1], stg_max_length=max_len, digit=digit), stamp_right) if(t[1]!='') else ''), tuple(stamps.items()))))
    elif((stamps!='') if(isinstance(stamps, str)) else False):
        stamp = '%s%s%s'%(stamp_left, stamps, stamp_right)
    elif((np.array(stamps).shape[0]>0) if(len(np.array(stamps).shape)>0) else False):
        stamps = [v for v in stamps if (not v in [''] if(isinstance(v, str)) else True)]
        stamp = adjoint_sep.join(list(map(lambda s:('%s%s%s'%(stamp_left, parse(s, stg_max_length=max_len, digit=digit), stamp_right)), stamps)))
    if(annih_when_stamps_empty and stamp==''):
        return ''
    stamp = outer_stamp_left + stamp + outer_stamp_right if(stamp!='') else ''
    stg = (stamp + stg if(location>0) else stg + stamp) if(location!=0) else stg
    stg = for_file_process(stg) if(for_file) else stg
    return stg


ALIGN_CENTER = 0
ALIGN_LEFT = -1
ALIGN_RIGHT = 1
def addlog(*log, max_len=None, level=1, function='log', stamps=None, click=None, click_anchor=None, click_stamp=None,
           log_counter=None, log_counter_stamp=None, log_counter_stamps=None, log_counter_ubd=5, log_counter_ubds={}, 
           reset_log_counter=False, reset_log_counter_value=0, log_when_unreset=False, colora='', encoding='utf-8',
           error_counter=None, max_logfile_error=10, dont_print=False, parse_digit=2, adjoint_sep=' ', handler=None, 
           fill_char=' ', fill_width=None, align=ALIGN_CENTER, **kwags):
    log = [parse(lg, parse_digit, stg_max_length=max_len) for lg in log if ((lg!='' and lg!='\n') if(isinstance(lg, str)) else True)]
    log_rawstg = stamp_process('',list(map(lambda x:x, log)) ,'' ,'' ,'' ,adjoint_sep, max_len=max_len)
    # 處理填充和對齊
    if fill_width is not None and isinstance(log_rawstg, str):
        if align == ALIGN_CENTER:
            log_rawstg = log_rawstg.center(fill_width, fill_char)
        elif align < 0:
            log_rawstg = log_rawstg.ljust(fill_width, fill_char)
        elif align > 0:
            log_rawstg = log_rawstg.rjust(fill_width, fill_char)
    stamps = stamps if(isinstance(stamps, list) or isinstance(stamps, dict)) else []
    kwags['annih_when_stamps_empty'] = kwags.get('annih_when_stamps_empty', False)
    log = stamp_process(log_rawstg, stamps=stamps, max_len=max_len, **kwags)
    if(log):
        if(log_when_unreset and isinstance(log_counter, dict)):
            if(isinstance(log_counter_stamps, (list, dict))):
                for lc_stamp in log_counter_stamps:
                    if(log_counter.get(lc_stamp, reset_log_counter_value)==reset_log_counter_value):
                        return
            if(log_counter.get(log_counter_stamp, reset_log_counter_value)==reset_log_counter_value):
                return
        if(not reset_log_counter and isinstance(log_counter_ubds, dict) and isinstance(log_counter, dict)):
            if(isinstance(log_counter_stamps, (list, dict))):
                for lc_stamp in log_counter_stamps:
                    ubd = log_counter_ubds.get(lc_stamp, log_counter_ubd)
                    if(log_counter.get(lc_stamp, 0)>ubd):
                        return
            ubd = log_counter_ubds.get(log_counter_stamp, log_counter_ubd)
            if(log_counter.get(log_counter_stamp, 0)>ubd):
                return
        display_max_columns_single = kwags.get('display_max_columns_single', '')
        display_max_rows_single = kwags.get('display_max_rows_single', '')
        if(display_max_columns_single!=''):
            pd.set_option('display.max_columns', display_max_columns_single)
        if(display_max_rows_single!=''):
            pd.set_option('display.max_rows', display_max_rows_single)
        if('longstgs' in kwags):
            longstgs = kwags['longstgs']
            longstgs = {s:200 for s in longstgs} if(type(longstgs)!=dict) else longstgs
            stgs_count = len(longstgs)
            format_count = log.count('%s')
            log = log%(*tuple([s[:longstgs[s]] for s in longstgs][:format_count]), 
                       *tuple(['%s']*(format_count - stgs_count)))
        #計時
        log = log[:max_len] if(isinstance(max_len, int)) else log
        dt_format = kwags.get('total_sec_digit', '%.2f(s)')
        if(isinstance(click, dict)):
            log_click = dcp(log)
            click_stamp = click_stamp if(isinstance(click_stamp, str)) else log_counter_stamp
            click_copy = click.copy()
            for k,v in click_copy.items():
                if(not isinstance(v, dt)):
                    continue
                dt_stg = '....%s費時%s'%(('[stamp:%s]'%k if(k!='') else ''), dt_format%((dt.now()-v).total_seconds()))
                log_click += dt_stg
                click[k] = dt.now()
            click.update({click_stamp:dt.now()}) if(not click_stamp in click) else None
        if(isinstance(click_anchor, dict)):
            log_click = dcp(log if(not isinstance(click, dict)) else log_click)
            click_stamp = click_stamp if(isinstance(click_stamp, str)) else log_counter_stamp
            for k,v in click_anchor.items():
                if(not isinstance(v, dt)):
                    continue
                dt_stg = '....%s費時%s'%(('[stamp:%s]'%k if(k!='') else ''), dt_format%((dt.now()-v).total_seconds()))
                log_click += dt_stg
            click_anchor.update({click_stamp:dt.now()}) if(not click_stamp in click_anchor) else None
        dont_print = False
        #存儲
        logfile = kwags.get('logfile', r'log\log.log')
        logOutput = dcp(log_click if(isinstance(click, dict) or isinstance(click_anchor, dict)) else log)
        print(colora + logOutput) if(not dont_print and logfile!=None and not isinstance(handler, str)) else None
        if(isinstance(handler, str)):
            setattr(handler, 'msgs', getattr(handler,'msgs',[]) + [logOutput])
        if(isinstance(log_counter, dict)):
            for log_counter_stamp_i in ((log_counter_stamps if(isinstance(log_counter_stamps, list)) else [])+[log_counter_stamp]):
                log_counter[log_counter_stamp_i] = log_counter.get(log_counter_stamp_i, 0) + 1
        if(isinstance_not_empty(logfile, str)):
            logfile = (logfile.replace('%t','%s'))%dt.now().strftime('%Y%m%d') if(
                                                logfile.find('%t')>-1) else logfile
            log_among_time = dcp(log_click if('click' in kwags or 'click_anchor' in kwags) else log)
            stg_now = dt.now().strftime('%Y-%m-%d %H:%M:%S\t')
            log_among_time = (stg_now + log_among_time + '\n') if (
                not log_among_time[-1]=='\n' if(len(log_among_time)>0) else True) else log_among_time
            try:
                error_counter = error_counter if(isinstance(error_counter, dict)) else {}
                with open(logfile ,'a', encoding=encoding) as f:
    	            f.write(log_among_time)
                error_counter.update({'logfile':0}) if('logfile' in error_counter) else None
            except:
                if(error_counter.get('logfile',0) < max_logfile_error):
                    print(colora + '[%d]logfile_error:%s.......'%(error_counter.get('logfile',0)+1, logfile))
                    error_counter['logfile'] = error_counter.get('logfile', 0) + 1
        if('display_max_columns_single' in kwags):
            pd.set_option('display.max_columns', 0)
        if('display_max_rows_single' in kwags):
            pd.set_option('display.max_rows', 0)
        if('abort_infrm' in kwags):
            abort_infrm = kwags['abort_infrm']
            i = len(abort_infrm)
            if(isinstance(stamps, dict)):
                abort_infrm[i] = dcp(stamps)
            elif(isiterable(stamps)):
                keys_dict = {i:stamps[i] for i in range(len(stamps))}
                abort_infrm[i] = {'key%d'%t[0]:'%s'%t[1] for t in keys_dict.items()}
            else:
                stamps = [stamps]
                keys_dict = {i:stamps[i] for i in range(len(stamps))}
                abort_infrm[i] = {'key%d'%t[0]:'%s'%t[1] for t in keys_dict.items()}
            abort_infrm[i]['msg'] = log_rawstg[:100]
    if(reset_log_counter and isinstance(log_counter, dict)):
        if(isinstance(log_counter_stamps, dict)):
            for k in log_counter_stamps: 
                log_counter.update({k:reset_log_counter_value}) if(k in log_counter and k!=None) else None
        else:
            if(log_counter_stamp==None):
                for k in log_counter: log_counter[k] = reset_log_counter_value
            else:
                log_counter.update({log_counter_stamp:reset_log_counter_value}) if(log_counter_stamp in log_counter) else None
        return


def addloger(**kwags):
    return lambda *s,**kws:addlog(*s,**kwags,**kws)


def execute(name, *containers, default=None, dominator=type(None), print_finded_label=False, print_value=False, 
            criterion=None, not_found_alarm=True, **specific_containers):
    try:
        old_default = dcp(default)
    except:
        old_default = default
    finded_label = None
    if(dominator!=type(None)):
        try:
            default = dcp(dominator)
            finded_label = '--dominated'
        except:
            default = dominator
    if(finded_label==None):
        for k,v in specific_containers.items():
            v = v if(isinstance(v, dict)) else {}
            if(name in v):
                default = v[name]
                finded_label = dcp(stamp_process('',['specific_containers',k],'','','',' '))
                break
    if(finded_label==None):
        for i,v in enumerate(containers):
            if(isinstance(v, dict)):
                if('m_%s'%name in v):
                    default = v['m_%s'%name]
                    finded_label = dcp(i)
                    break
                elif(name in v):
                    default = v[name]
                    finded_label = dcp(i)
                    break
            elif(hasattr(v, name)):
                default = getattr(v, name)
                finded_label = dcp(i)
                break
    print(finded_label) if(print_finded_label) else None
    (addlog('`%s`'%(str(default)[:200]), stamps=[name]) if(
        not isiterable(default)) else show_vector(default, stamps=[name])) if(print_value) else None
    if(finded_label==None):
        if(not_found_alarm):
            print('`%s` execute error:\n`%s`'%(name, str(default)[:200]))
            sys.exit(1)
        return default
    if(criterion!=None):
        if(not criterion(default)):
            if(not_found_alarm):
                print('`%s` execute error with criterion:\n`%s`'%(name, str(default)[:200]))
                sys.exit(1)
            return old_default
    specific_containers.get('project_buffer',{}).update({name:default})
    return default

def executeEasy(name, *args, default=None, not_found_alarm=False, **kwags):
    return execute('addlog', *args, default=default, not_found_alarm=not_found_alarm)


def exception_process(e, logfile=os.path.join('log_%t.txt'), stamps=None, max_len_stack=200, handler=None, colora=FAIL, **kwags):
    stamps = stamps if(isinstance(stamps, list)) else []
    exc_type, exc_obj, ex_stack = sys.exc_info()
    ex_stamps = {'lineno':'%d'%e.__traceback__.tb_lineno,
                 'name':'%s'%e.__traceback__.tb_frame.f_code.co_name,
                 'type':'%s'%exc_type}
    msg = '[ERROR]%s\n%s%s\n%s'%(e.__traceback__.tb_frame.f_code.co_filename,
        stamp_process('',ex_stamps), stamp_process('', stamps), '錯誤訊息:\n%s\n'%str(e))
    None if(isinstance(handler,str)) else addlog('--------------------------------------------------', logfile=logfile, colora=colora)
    traceback_things = traceback.extract_tb(ex_stack)
    for stack in traceback.extract_tb(ex_stack):
        addlog(str(stack)[:max_len_stack], logfile=logfile, handler=handler, colora=colora, **kwags)
    None if(isinstance(handler,str)) else addlog('--------------------------------------------------', logfile=logfile, colora=colora)
    kwags.update({'annih_when_stamps_empty':False})
    addlog(msg, logfile=logfile, handler=handler, colora=colora, **kwags)


def detect_txt_encoding(file, encoding_default=None, chardetCount_ubd=5000):
    if(not isinstance(encoding_default, str)):
        try:
            with open(file, 'rb') as f:
                raw_data = f.read(chardetCount_ubd)  # 读取一部分文件内容用于编码检测
                resault = chd.detect(raw_data)
            return resault['encoding']
        except Exception as e:
            exception_process(e,logfile='')
            return encoding_default
    return encoding_default


def load_json(file, chardetCount_ubd=5000, encoding=None, **kwags):
    if(not os.path.exists(file)):
        return {}
    encodingDefault = detect_txt_encoding(file, chardetCount_ubd=chardetCount_ubd)
    encoding = encoding if(isinstance_not_empty(encoding, str)) else encodingDefault
    with open(file, encoding=encoding) as f:
        data = json.load(f)
        f.close()
    return data


def _safe_print(text: str, end: str = "\n") -> None:
    """
    Print text without crashing on console encoding issues (e.g. cp950).
    """
    try:
        print(text, end=end)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        fallback = text.encode(enc, errors="replace").decode(enc, errors="replace")
        print(fallback, end=end)


def _stamps_to_text(stamps: Any) -> str:
    if stamps is None:
        return ""
    if isinstance(stamps, dict):
        parts = [f"{k}:{v}" for k, v in stamps.items()]
    elif isinstance(stamps, (list, tuple, set)):
        parts = [str(s) for s in stamps]
    else:
        parts = [str(stamps)]
    return "".join(f"[{p}]" for p in parts if p != "")


def _append_log_line(logfile: str, line: str, encoding: str = "utf-8") -> None:
    if not logfile:
        return
    _ensure_outputs_root()
    log_path = _to_output_path(logfile)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with open(log_path, "a", encoding=encoding, newline="\n") as f:
        f.write(line)
        if not line.endswith("\n"):
            f.write("\n")


def interval_overlap_seconds(
    a0: Any,
    a1: Any,
    b0: Any,
    b1: Any,
) -> Optional[float]:
    """Return overlap seconds of intervals [a0,a1] and [b0,b1]; None if any endpoint is not datetime."""
    if not all(isinstance(x, dt) for x in (a0, a1, b0, b1)):
        return None
    lo = max(a0.timestamp(), b0.timestamp())
    hi = min(a1.timestamp(), b1.timestamp())
    return max(0.0, hi - lo)


class _NumpyJSONEncoder(json.JSONEncoder):
    def default(self, obj: Any) -> Any:
        if np is not None:
            if isinstance(obj, np.integer):
                return int(obj)
            if isinstance(obj, np.floating):
                return float(obj)
            if isinstance(obj, np.ndarray):
                return obj.tolist()
        if hasattr(obj, "tolist") and callable(getattr(obj, "tolist", None)):
            try:
                return obj.tolist()
            except Exception:
                pass
        if hasattr(obj, "isoformat") and callable(getattr(obj, "isoformat", None)):
            try:
                return obj.isoformat()
            except Exception:
                pass
        return super().default(obj)


class LOGger:
    """
    Minimal compatibility layer for legacy LOGger usage in tmp2/dataProcess.
    """

    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"

    dcp = staticmethod(dcp)

    @staticmethod
    def addloger(**kwargs: Any) -> Callable[..., None]:
        base_logfile = kwargs.get("logfile", "")
        default_colora = kwargs.get("colora", "")

        def _logger(*logs: Any, **kws: Any) -> None:
            stamps = kws.get("stamps")
            colora = kws.get("colora", default_colora)
            logfile = kws.get("logfile", base_logfile)
            end = kws.get("end", "\n")
            show_time = kws.get("show_time", True)

            message = " ".join(str(x) for x in logs)
            stamp_text = _stamps_to_text(stamps)
            prefix = dt.now().strftime("%Y-%m-%d %H:%M:%S ") if show_time else ""
            line = f"{prefix}{stamp_text}{message}"

            if colora:
                _safe_print(f"{colora}{line}{LOGger.ENDC}", end=end)
            else:
                _safe_print(line, end=end)

            if isinstance(logfile, str) and logfile.strip():
                _append_log_line(logfile, line)

        return _logger

    @staticmethod
    def exception_process(
        e: Exception,
        logfile: str = "log_%t.txt",
        stamps: Any = None,
        max_len_stack: int = 200,
        handler: Optional[Callable[..., Any]] = None,
        colora: str = FAIL,
        **kwargs: Any,
    ) -> None:
        stamp_text = _stamps_to_text(stamps)
        tb_lines = traceback.format_exception(type(e), e, e.__traceback__)
        head = f"[ERROR]{stamp_text} {e}"

        if handler is not None:
            try:
                handler(head)
                for line in tb_lines:
                    handler(line[:max_len_stack])
                return
            except Exception:
                pass

        if colora:
            _safe_print(f"{colora}{head}{LOGger.ENDC}")
        else:
            _safe_print(head)

        for line in tb_lines:
            _safe_print(line[:max_len_stack].rstrip("\n"))

        if isinstance(logfile, str) and logfile.strip():
            _append_log_line(logfile, head)
            for line in tb_lines:
                _append_log_line(logfile, line[:max_len_stack].rstrip("\n"))

    @staticmethod
    def CreateFile(
        path: Any,
        method: Callable[[str], Any],
        maxpathlen: int = 10,
        retry_ubd: Optional[int] = None,
        retry_slt_method: Callable[[], float] = lambda: 0.1,
        rename_if_retry: Optional[bool] = None,
    ) -> tuple[str, str]:
        del maxpathlen  # compatibility argument
        target_base = _to_output_path(path)
        attempts = int(retry_ubd if retry_ubd is not None else 1) + 1
        rename_retry = bool(rename_if_retry)
        last_exc: Optional[Exception] = None

        for i in range(attempts):
            target = target_base
            if i > 0 and rename_retry:
                target = target_base.with_name(f"{target_base.stem}_RETRY{i}{target_base.suffix}")
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                method(str(target))
                return str(target.parent), target.name
            except Exception as exc:  # pragma: no cover
                last_exc = exc
                if i < attempts - 1:
                    try:
                        time.sleep(float(retry_slt_method()))
                    except Exception:
                        time.sleep(0.1)

        if last_exc is not None:
            LOGger.exception_process(last_exc, stamps=["CreateFile", str(target_base)])
        return str(target_base.parent), target_base.name

    @staticmethod
    def save_json(
        data: Any,
        file: Any,
        mode: str = "w",
        indent: int = 4,
        data_sieve_method: Optional[Callable[[Any], None]] = None,
        ensure_ascii: bool = False,
        encoding: str = "utf-8",
    ) -> bool:
        try:
            if data_sieve_method is not None:
                data_sieve_method(data)
            target = _to_output_path(file)
            target.parent.mkdir(parents=True, exist_ok=True)
            with open(target, mode, encoding=encoding, newline="\n") as f:
                json.dump(
                    data,
                    f,
                    indent=indent,
                    cls=_NumpyJSONEncoder,
                    ensure_ascii=ensure_ascii,
                )
            return True
        except Exception as exc:  # pragma: no cover
            LOGger.exception_process(exc, stamps=["save_json", str(file)])
            return False


def isnonnumber(x: Any) -> bool:
    """Compatibility helper from dataframeprocedure: True when value is not numeric-castable."""
    try:
        float(x)
        return False
    except Exception:
        return True


def isiterable(a: Any, exceptions: tuple[type, ...] = (str, dict), type_stg_exceptions: tuple[str, ...] = ()) -> bool:
    """Compatibility helper from dataframeprocedure/LOGger."""
    if isinstance(a, exceptions):
        return False
    if any(exc in str(type(a)) for exc in type_stg_exceptions):
        return False
    try:
        iter(a)
        return True
    except Exception:
        return False


def normalize_idx(expr: Any) -> str:
    """Normalize $(r,c) / $r tokens into explicit 2D $(r,c) form."""
    import re

    s = str(expr)
    s = re.sub(r"(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", r"$(\1,\2)", s)
    return re.sub(r"\$(\d+)", lambda m: f"$({m.group(1)},0)", s)


def repl_names(expr: Any, name_to_index: dict[str, int]) -> str:
    """Replace row-name tokens in expression with $(row_index,0)."""
    x = str(expr)
    for name in sorted(name_to_index.keys(), key=len, reverse=True):
        x = x.replace(name, f"$({name_to_index[name]},0)")
    return x
