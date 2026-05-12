import os
import json
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--root_path', type=str, default=None)
parser.add_argument('--root_paths', type=str, nargs='+', default=None,
                    help='直接指定多个 root path，按顺序作为 trial 0, 1, 2, ...')
parser.add_argument('--meta_path', type=str, default='./evaluaton_data/test_oswmcp_feasible.json')
parser.add_argument('--trials', type=int, nargs='+', default=[0, 1, 2])
parser.add_argument('--detail', action='store_true')
parser.add_argument('--filter_fail', action='store_true')
parser.add_argument('--aggregate', action='store_true',
                    help='Aggregate domains into multi_apps vs other')
parser.add_argument('--tool_beneficial_path', type=str,
                    default='./evaluation_data/has_tool_feasible.json')
parser.add_argument('--no_tool_beneficial_path', type=str,
                    default='./evaluation_data/no_tool_feasible.json')
parser.add_argument('--eval_15', action='store_true')
parser.add_argument('--eval_30', action='store_true')

args = parser.parse_args()


if args.eval_15:
    TARGET_RESULT_FILE_PATH = "result_15.txt"
elif args.eval_30:
    TARGET_RESULT_FILE_PATH = "result_30.txt"
else: 
    TARGET_RESULT_FILE_PATH = "result.txt"
    


# ---------- 校验参数 ----------
if args.root_paths is not None and args.root_path is not None:
    raise ValueError('--root_path 和 --root_paths 不能同时指定')
if args.root_paths is None and args.root_path is None:
    raise ValueError('必须指定 --root_path 或 --root_paths 之一')

# ---------- 统一成 trial_id -> root_path 的映射 ----------
# 模式A：root_path + trials  -> subdir = root_path / trial_id / domain / task_id
# 模式B：root_paths           -> subdir = root_paths[i] / domain / task_id，trial_id = i
USE_MULTI_ROOT = args.root_paths is not None

if USE_MULTI_ROOT:
    trial_root_map = {i: p for i, p in enumerate(args.root_paths)}
    trials = list(trial_root_map.keys())
    print(f'模式B: 使用 root_paths，共 {len(trials)} 个 trial')
    for i, p in trial_root_map.items():
        print(f'  trial {i}: {p}')
else:
    root_path = args.root_path
    trials = args.trials
    if len(trials) == 0:
        trials = sorted([int(x) for x in os.listdir(root_path)])
    trial_root_map = {i: root_path for i in trials}
    print(f'模式A: 使用 root_path={root_path}，trials={trials}')

print('trials:', trials)


def get_subdir(trial_id, domain, task_id):
    """根据模式返回正确的子目录路径"""
    if USE_MULTI_ROOT:
        return os.path.join(trial_root_map[trial_id], domain, task_id)
    else:
        return os.path.join(trial_root_map[trial_id], str(trial_id), domain, task_id)


with open(args.meta_path, 'r') as f:
    data = json.load(f)

filter_fail = args.filter_fail


# ---------- 读取 tool_beneficial / non_tool_beneficial 任务集合 ----------
def load_task_set(path):
    with open(path, 'r') as f:
        raw = json.load(f)
    task_set = set()
    if isinstance(raw, dict):
        for domain, task_list in raw.items():
            for task_id in task_list:
                task_set.add((domain, task_id))
    elif isinstance(raw, list):
        for item in raw:
            if isinstance(item, dict):
                task_set.add((item['domain'], item['task_id']))
            elif isinstance(item, str):
                parts = item.split('/', 1)
                if len(parts) == 2:
                    task_set.add((parts[0], parts[1]))
    return task_set


tool_beneficial_set    = load_task_set(args.tool_beneficial_path)
no_tool_beneficial_set = load_task_set(args.no_tool_beneficial_path)


# ---------- traj 解析工具函数 ----------
def load_traj(subdir):
    traj_path = os.path.join(subdir, 'traj.jsonl')
    if not os.path.exists(traj_path):
        return None
    with open(traj_path, 'r') as f:
        lines = [l for l in f.readlines() if l.strip()]
    return [json.loads(l) for l in lines]


def parse_traj_info(traj):
    if traj is None:
        return False, 0, None

    effective  = traj[1:]
    tool_calls = sum(1 for step in effective if isinstance(step.get('action'), dict))
    used_tool  = tool_calls > 0

    step_num = None
    if traj:
        last = traj[-1]
        if 'step_num' in last:
            step_num = last['step_num']

    return used_tool, tool_calls, step_num


# ---------- load_result ----------
def load_result(subdir):
    if os.path.exists(os.path.join(subdir, TARGET_RESULT_FILE_PATH)):
        result_path = os.path.join(subdir, TARGET_RESULT_FILE_PATH)
    else: 
        result_path = os.path.join(subdir, 'result.txt')
    
    traj_path   = os.path.join(subdir, 'traj.jsonl')
    if not (os.path.exists(traj_path) and os.path.exists(result_path)):
        return None

    with open(result_path, 'r') as f:
        result = f.read().strip()
    if result.lower() == 'false':
        result = 0
    elif result.lower() == 'true':
        result = 1
    else:
        result = float(result)

    if filter_fail:
        traj = load_traj(subdir)
        if traj and traj[-1].get('action') == 'FAIL' and len(traj) >= 16:
            result = 0

    return result


total_num = sum(len(datalist) for datalist in data.values())

# ---------- 聚合每个任务的多轮结果 ----------
acc = {}
invalid_tasks = {}  # 新增: trial_id -> list of (domain, task_id, subdir)
for domain, datalist in data.items():
    for d in datalist:
        for i in trials:
            subdir = get_subdir(i, domain, d)
            r = load_result(subdir)
            if r is None:
                # ← 收集 invalid 信息
                if i not in invalid_tasks:
                    invalid_tasks[i] = []
                invalid_tasks[i].append((domain, d, subdir))
                continue
            traj = load_traj(subdir)
            used_tool, tool_calls, step_num = parse_traj_info(traj)

            key   = f'{domain}_{d}'
            entry = {
                'trial':      i,
                'result':     r,
                'used_tool':  used_tool,
                'tool_calls': tool_calls,
                'step_num':   step_num,
            }
            if key not in acc:
                acc[key] = [entry]
            else:
                acc[key].append(entry)

# ---------- 打印 invalid tasks ----------
print('\n' + '=' * 60)
print('[Invalid Tasks per Trial]')
print('=' * 60)
for trial_id in trials:
    inv_list = invalid_tasks.get(trial_id, [])
    print(f'\nTrial {trial_id}: {len(inv_list)} invalid tasks')
    for domain, task_id, subdir in inv_list:
        # 诊断原因
        traj_exists   = os.path.exists(os.path.join(subdir, 'traj.jsonl'))
        result_exists = os.path.exists(os.path.join(subdir, 'result.txt'))
        reason = []
        if not traj_exists:
            reason.append('missing traj.jsonl')
        if not result_exists:
            reason.append('missing result.txt')
        if not os.path.exists(subdir):
            reason = ['subdir not exists']
        print(f'  {domain}_{task_id}  [{", ".join(reason)}]  ({subdir})')


# ---------- 整体 pass@k ----------
acc_list  = [any(t['result'] >= 0.5 for t in x) for x in acc.values()]
final_acc = sum(acc_list) / len(acc_list)
print(f'Pass@{len(trials)} Acc: {final_acc}, success: {sum(acc_list)}, valid task: {len(acc_list)}/{total_num}')


# ---------- detail ----------
if args.detail:
    output_text = ''

    for domain, datalist in data.items():
        for task_id in datalist:
            if f'{domain}_{task_id}' not in acc:
                output_text += f'{domain}\t{task_id}\t{0}\n'
                continue
            cur_acc     = any(t['result'] >= 0.5 for t in acc[f'{domain}_{task_id}'])
            num_success = sum(1 for t in acc[f'{domain}_{task_id}'] if t['result'] >= 0.5)
            num_trials  = len(acc[f'{domain}_{task_id}'])
            if cur_acc:
                with open(f'./evaluation_examples/examples/{domain}/{task_id}.json', 'r') as f:
                    task_config = json.load(f)
                    instruction = task_config['instruction']

                lengthes = []
                for trial_id in trials:
                    sub_dir   = get_subdir(trial_id, domain, task_id)
                    traj      = load_traj(sub_dir)
                    if traj is None:
                        continue
                    trial_acc = load_result(sub_dir)
                    if trial_acc is None:
                        continue
                    traj_len = len(traj)
                    if trial_acc >= 0.5:
                        lengthes.append(str(traj_len) + '(✓)')
                    else:
                        lengthes.append(str(traj_len))

                print(f'{domain}_{task_id} is successful {num_success}/{num_trials} '
                      f'|traj length: [{", ".join(lengthes)}]: {instruction}')

            output_text += f'{domain}\t{task_id}\t{num_success}\n'

    with open('result.txt', 'w') as f:
        f.write(output_text)


# ---------- per-trial avg ----------
def show_acc_per_trial(trials, data, acc):
    num_valid = sum(len(datalist) for datalist in data.values())
    mean_acc  = []
    for trial_id in trials:
        acc_per_trial = []
        for domain, datalist in data.items():
            for task_id in datalist:
                if f"{domain}_{task_id}" not in acc:
                    continue
                for result in acc[f"{domain}_{task_id}"]:
                    if result['trial'] == trial_id:
                        acc_per_trial.append(1.0 if result['result'] >= 0.5 else 0.0)

        acc_per_trial_ = sum(acc_per_trial) / num_valid
        mean_acc.append(acc_per_trial_)
        print(f'Trial {trial_id} acc: {acc_per_trial_:.4f}, num valid: {len(acc_per_trial)}/{num_valid}')

    print(f'Mean acc: {sum(mean_acc) / len(mean_acc):.4f}, '
          f'success: {round(sum(mean_acc) / len(mean_acc) * num_valid)}/{num_valid}')


# ---------- pass@k + avg@k，支持 aggregate ----------
def show_acc_by_group(data, acc, trials, aggregate=False):
    if aggregate:
        group_keys = {'multi_apps': [], 'other': []}
        for domain, datalist in data.items():
            gk = 'multi_apps' if domain == 'multi_apps' else 'other'
            for task_id in datalist:
                group_keys[gk].append((domain, task_id))
    else:
        group_keys = {}
        for domain, datalist in data.items():
            group_keys[domain] = [(domain, task_id) for task_id in datalist]

    all_passk = []
    all_avgk  = []

    for group_name, task_list in group_keys.items():
        passk_list = []
        avgk_list  = []

        for domain, task_id in task_list:
            key = f'{domain}_{task_id}'
            if key not in acc:
                passk_list.append(False)
                avgk_list.append(0.0)
            else:
                results = [t['result'] for t in acc[key]]
                passk_list.append(any(r >= 0.5 for r in results))
                avgk_list.append(sum(1.0 if r >= 0.5 else 0.0 for r in results) / len(trials))

        total  = len(task_list)
        n_pass = sum(passk_list)
        passk  = n_pass / total if total > 0 else 0.0
        avgk   = sum(avgk_list) / total if total > 0 else 0.0

        print(f'[{group_name}]'
              f'  Pass@{len(trials)}: {passk:.4f} ({n_pass}/{total})'
              f'  Avg@{len(trials)}: {avgk:.4f}')

        all_passk.extend(passk_list)
        all_avgk.extend(avgk_list)

    total  = len(all_passk)
    n_pass = sum(all_passk)
    print(f'[overall]'
          f'  Pass@{len(trials)}: {n_pass/total:.4f} ({n_pass}/{total})'
          f'  Avg@{len(trials)}: {sum(all_avgk)/total:.4f}')


# ---------- 调用已有指标 ----------
print()
show_acc_per_trial(trials, data, acc)

print()
show_acc_by_group(data, acc, trials, aggregate=args.aggregate)


# ----------为了paper进行visualization
# ---------- 按指定顺序打印 avg@k acc list ----------
def show_ordered_avgk_list(data, acc, trials):
    category_order = ['chrome', 'gimp', 'libreoffice_calc', 'libreoffice_impress', 'libreoffice_writer',
                      'os', 'thunderbird', 'vlc', 'vs_code', 'multi_apps', 'overall']

    # domain 名称到 category_order 中 key 的映射（如有别名可在此处理）
    domain_alias = {
        'vscode': 'vs code',
        'vs code': 'vs code',
        'multi_apps': 'multi_apps',
        'multiple apps': 'multi_apps',
    }

    # 先按 category 聚合 avgk
    category_avgk = {}   # category_key -> list of per-task avg@k values
    overall_avgk  = []

    for domain, datalist in data.items():
        cat_key = domain_alias.get(domain, domain)  # 尝试别名映射，否则用原名
        for task_id in datalist:
            key = f'{domain}_{task_id}'
            if key not in acc:
                val = 0.0
            else:
                results = [t['result'] for t in acc[key]]
                val = sum(1.0 if r >= 0.5 else 0.0 for r in results) / len(trials)

            if cat_key not in category_avgk:
                category_avgk[cat_key] = []
            category_avgk[cat_key].append(val)
            overall_avgk.append(val)

    # 计算每个 category 的均值
    result_line = []
    for cat in category_order:
        if cat == 'overall':
            val = sum(overall_avgk) / len(overall_avgk) if overall_avgk else 0.0
        else:
            vals = category_avgk.get(cat, [])
            val  = sum(vals) / len(vals) if vals else 0.0
        result_line.append(f'{val * 100:.1f}')

    print()
    print('[Ordered Avg@{} Acc List]'.format(len(trials)))
    print('Order: ' + ', '.join(category_order))
    print(', '.join(result_line))

show_ordered_avgk_list(data, acc, trials)



# ---------- subtask metrics ----------
def compute_subtask_metrics(data, acc, trials, task_set, group_name):
    task_list = []
    for domain, datalist in data.items():
        for task_id in datalist:
            if (domain, task_id) in task_set:
                task_list.append((domain, task_id))

    total = len(task_list)
    if total == 0:
        print(f'[{group_name}] No tasks found.')
        return None

    is_tool_beneficial = (group_name == 'tool_beneficial')

    avgk_list             = []
    tir_list              = []
    tool_calling_list     = []
    steps_list            = []
    completion_steps_list = []

    for domain, task_id in task_list:
        key = f'{domain}_{task_id}'

        if key not in acc:
            avgk_list.append(0.0)
            tir_list.append(0.0)
            continue

        entries = acc[key]

        # ---------- avg@k accuracy ----------
        n_success = sum(1 for e in entries if e['result'] >= 0.5)
        avgk_list.append(n_success / len(trials))

        # ---------- TIR ----------
        # tool_beneficial:     used_tool and success
        # non_tool_beneficial: not used_tool and success
        n_tir = 0
        for e in entries:
            success   = e['result'] >= 0.5
            used_tool = e['used_tool']
            if is_tool_beneficial:
                if used_tool and success:
                    n_tir += 1
            else:
                if (not used_tool) and success:
                    n_tir += 1
        tir_list.append(n_tir / len(trials))

        # ---------- avg_tool_calling / avg_steps / avg_completion_steps ----------
        for e in entries:
            tool_calling_list.append(e['tool_calls'])
            if e['step_num'] is not None:
                steps_list.append(e['step_num'])
            if e['result'] >= 0.5 and e['step_num'] is not None:
                completion_steps_list.append(e['step_num'])

    avgk                 = sum(avgk_list) / total
    tir                  = sum(tir_list) / total
    avg_tool_calling     = sum(tool_calling_list) / len(tool_calling_list) if tool_calling_list else 0.0
    avg_steps            = sum(steps_list) / len(steps_list) if steps_list else 0.0
    avg_completion_steps = sum(completion_steps_list) / len(completion_steps_list) if completion_steps_list else 0.0

    print(f'[{group_name}]  total={total}')
    print(f'  Avg@{len(trials)} Acc (%)            : {avgk * 100:.2f}')
    print(f'  Avg@{len(trials)} TIR (%)             : {tir * 100:.2f}')
    print(f'  Avg Tool Calling (per trial) : {avg_tool_calling:.4f}')
    print(f'  Avg Steps (per trial)        : {avg_steps:.2f}')
    print(f'  Avg Completion Steps         : {avg_completion_steps:.2f}')

    return {
        'total':                 total,
        'avgk_list':             avgk_list,
        'tir_list':              tir_list,
        'tool_calling_list':     tool_calling_list,
        'steps_list':            steps_list,
        'completion_steps_list': completion_steps_list,
    }


def show_tool_beneficial_metrics(data, acc, trials, tool_beneficial_set, no_tool_beneficial_set):
    print('=' * 60)
    print(f'[Tool-Beneficial Analysis]  Avg@{len(trials)} only')
    print('=' * 60)

    tb_result  = compute_subtask_metrics(data, acc, trials, tool_beneficial_set,    'tool_beneficial')
    ntb_result = compute_subtask_metrics(data, acc, trials, no_tool_beneficial_set, 'non_tool_beneficial')

    print()
    print('[overall (tool_beneficial + non_tool_beneficial)]')

    if tb_result is None and ntb_result is None:
        print('  No data available.')
        return

    all_avgk         = []
    all_tir          = []
    all_tool_calling = []
    all_steps        = []
    all_completion   = []

    for res in [tb_result, ntb_result]:
        if res is None:
            continue
        all_avgk.extend(res['avgk_list'])
        all_tir.extend(res['tir_list'])
        all_tool_calling.extend(res['tool_calling_list'])
        all_steps.extend(res['steps_list'])
        all_completion.extend(res['completion_steps_list'])

    total                = len(all_avgk)
    avg_tool_calling     = sum(all_tool_calling) / len(all_tool_calling) if all_tool_calling else 0.0
    avg_steps            = sum(all_steps) / len(all_steps) if all_steps else 0.0
    avg_completion_steps = sum(all_completion) / len(all_completion) if all_completion else 0.0

    print(f'  total={total}')
    print(f'  Avg@{len(trials)} Acc (%)            : {sum(all_avgk) / total * 100:.2f}')
    print(f'  Avg@{len(trials)} TIR (%)           : {sum(all_tir) / total * 100:.2f}')
    print(f'  Avg Tool Calling (per trial) : {avg_tool_calling:.4f}')
    print(f'  Avg Steps (per trial)        : {avg_steps:.2f}')
    print(f'  Avg Completion Steps         : {avg_completion_steps:.2f}')


# ---------- 调用新增指标 ----------
print()
show_tool_beneficial_metrics(data, acc, trials, tool_beneficial_set, no_tool_beneficial_set)
