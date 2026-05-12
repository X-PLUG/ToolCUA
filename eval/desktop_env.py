from __future__ import annotations

import logging
import os
import time
import ast
import math
import re
import json
from typing import Callable, Any, Optional, Tuple
from typing import List, Dict, Union

import gymnasium as gym

from desktop_env.controllers.python import PythonController
from desktop_env.controllers.setup import SetupController
from desktop_env.evaluators import metrics, getters
from desktop_env.providers import create_vm_manager_and_provider

logger = logging.getLogger("desktopenv.env")

Metric = Callable[[Any, Any], float]
Getter = Callable[[gym.Env, Dict[str, Any]], Any]

MAX_RETRIES = 5 # Maximum retries for environment setup

# import asyncio
# from fastmcp import Client


import os
current_dir = os.path.dirname(os.path.abspath(__file__))
STATIC_MCP_DIR = os.path.join(os.path.dirname(current_dir), "GUI", "updated_mcp_server_clean", "mcp_server")
STATIC_MCP_CLIENT_PATH = os.path.join(os.path.dirname(current_dir), "GUI", "updated_mcp_server_clean", "client.py")
REMOTE_MCP_SERVER_LOG_PATH = "/tmp/osworld_mcp_server.log"



def _fix_pyautogui_less_than_bug(command: str) -> str:
    """
    Fix PyAutoGUI '<' character bug by converting it to hotkey("shift", ',') calls.
    
    This fixes the known PyAutoGUI issue where typing '<' produces '>' instead.
    References:
    - https://github.com/asweigart/pyautogui/issues/198
    - https://github.com/xlang-ai/OSWorld/issues/257
    
    Args:
        command (str): The original pyautogui command
        
    Returns:
        str: The fixed command with '<' characters handled properly
    """
    # Pattern to match press('<') or press('\u003c') calls  
    press_pattern = r'pyautogui\.press\(["\'](?:<|\\u003c)["\']\)'

    # Handle press('<') calls
    def replace_press_less_than(match):
        return 'pyautogui.hotkey("shift", ",")'
    
    # First handle press('<') calls
    command = re.sub(press_pattern, replace_press_less_than, command)

    # Pattern to match typewrite calls with quoted strings
    typewrite_pattern = r'pyautogui\.typewrite\((["\'])(.*?)\1\)'
    
    # Then handle typewrite calls
    def process_typewrite_match(match):
        quote_char = match.group(1)
        content = match.group(2)
        
        # Preprocess: Try to decode Unicode escapes like \u003c to actual '<'
        # This handles cases where '<' is represented as escaped Unicode
        try:
            # Attempt to decode unicode escapes
            decoded_content = content.encode('utf-8').decode('unicode_escape')
            content = decoded_content
        except UnicodeDecodeError:
            # If decoding fails, proceed with original content to avoid breaking existing logic
            pass  # English comment: Graceful degradation - fall back to original content if decoding fails
        
        # Check if content contains '<'
        if '<' not in content:
            return match.group(0)
        
        # Split by '<' and rebuild
        parts = content.split('<')
        result_parts = []
        
        for i, part in enumerate(parts):
            if i == 0:
                # First part
                if part:
                    result_parts.append(f"pyautogui.typewrite({quote_char}{part}{quote_char})")
            else:
                # Add hotkey for '<' and then typewrite for the rest
                result_parts.append('pyautogui.hotkey("shift", ",")')
                if part:
                    result_parts.append(f"pyautogui.typewrite({quote_char}{part}{quote_char})")
        
        return '; '.join(result_parts)
    
    command = re.sub(typewrite_pattern, process_typewrite_match, command)
    
    return command


class DesktopEnv(gym.Env):
    """
    DesktopEnv with OpenAI Gym interface. It provides a desktop environment for setting and evaluating desktop automation tasks.
    """
    def __init__(
            self,
            provider_name: str = "vmware",
            region: str = None,
            path_to_vm: str = None,
            snapshot_name: str = "init_state",
            action_space: str = "mcp",
            cache_dir: str = "cache",
            screen_size: Tuple[int] = (int(os.environ.get("SCREEN_WIDTH", 1920)), int(os.environ.get("SCREEN_HEIGHT", 1080))),
            headless: bool = False,
            require_a11y_tree: bool = True,
            require_terminal: bool = False,
            os_type: str = "Ubuntu",
            enable_proxy: bool = False,
            client_password: str = "",
            shuffle: bool = False,
            rag: bool = True,
            order_param: bool = False
    ):
        """
        Args:
            provider_name (str): virtualization provider name, default to "vmware"
            region (str): the region for allocate machines, work for cloud services, default to  "us-east-1"
            path_to_vm (str): path to .vmx file
            snapshot_name (str): snapshot name to revert to, default to "init_state"
            action_space (str): "computer_13" | "pyautogui"
            cache_dir (str): cache directory to cache task-related stuffs like
              reference file for evaluation
            screen_size (Tuple[int]): screen size of the VM
            headless (bool): whether to run the VM in headless mode
            require_a11y_tree (bool): whether to require accessibility tree
            require_terminal (bool): whether to require terminal output
            os_type (str): operating system type, default to "Ubuntu"
            enable_proxy (bool): whether to enable proxy support, default to False
        """
        # Initialize VM manager and vitualization provider
        self.order_param = order_param
        self.rag = rag
        self.shuffle = shuffle
        self.region = region
        self.provider_name = provider_name
        self.enable_proxy = enable_proxy  # Store proxy enablement setting
        if client_password == "":
            if self.provider_name == "aws":
                self.client_password = "osworld-public-evaluation"
            else:
                self.client_password = "password"
        else:
            self.client_password = client_password

        self.screen_width = screen_size[0]
        self.screen_height = screen_size[1]

        # Default 
        self.server_port = 5000
        self.chromium_port = 9222
        self.vnc_port = 8006
        self.vlc_port = 8080
        self.mcp_port = 9292
        
        # Initialize with default (no proxy) provider
        self.current_use_proxy = False
        self.manager, self.provider = create_vm_manager_and_provider(provider_name, region, use_proxy=False, )

        self.os_type = os_type

        # Track whether environment has been used (step/setup) to optimize snapshot revert
        # docker, aws, gcp, azure are always unused as the emulator starts from a clean state
        # vmware, virtualbox are always used as the emulator starts from a dirty state
        if self.provider_name in {"docker", "aws", "gcp", "azure", "aliyun", "volcengine"}:
            self.is_environment_used = False
        elif self.provider_name in {"vmware", "virtualbox"}:
            self.is_environment_used = True
        else:
            raise ValueError(f"Invalid provider name: {self.provider_name}")

        # Initialize environment variables
        if path_to_vm:
            self.path_to_vm = os.path.abspath(os.path.expandvars(os.path.expanduser(path_to_vm))) \
                if provider_name in {"vmware", "virtualbox"} else path_to_vm
        else:
            self.path_to_vm = self.manager.get_vm_path(os_type=self.os_type, region=region, screen_size=(self.screen_width, self.screen_height))
        
        try:
            self.snapshot_name = snapshot_name
            self.cache_dir_base: str = cache_dir
            # todo: add the logic to get the screen size from the VM
            self.headless = headless
            self.require_a11y_tree = require_a11y_tree
            self.require_terminal = require_terminal

            assert action_space in ["computer_13", "pyautogui", "claude_computer_use", "autoglm_computer_use", "mcp"]
            self.action_space = action_space  # todo: refactor it to the ActType

            # Initialize emulator and controller
            logger.info("Initializing...")
            self._start_emulator()

            # mode: human or machine
            self.instruction = None

            # episodic stuffs, like counters, will be updated or reset
            # when calling self.reset()
            self._traj_no: int = -1
            self._step_no: int = 0
            self.action_history: List[Dict[str, any]] = []
        except Exception as e:
            logger.error(f"Failed to initialize DesktopEnv: {e}")
            # If initialization fails, we should clean up the VM
            try:
                self.close()
                self.manager.delete_vm(self.path_to_vm, self.region)
                logger.info(f"Cleaned up VM {self.path_to_vm}.")
            except Exception as cleanup_error:
                logger.error(f"Failed to clean up VM {self.path_to_vm}: {cleanup_error}")
            raise

    def _start_emulator(self, docker_overlap=False):
        docker_overlap = (self.action_space == 'mcp')        

        # Power on the virtual machine
        self.provider.start_emulator(self.path_to_vm, self.headless, self.os_type, docker_overlap)

        # Get the ip from the virtual machine, and setup the controller
        vm_ip_ports = self.provider.get_ip_address(self.path_to_vm).split(':')
        self.vm_ip = vm_ip_ports[0]
        # Get the ports from the virtual machine (for Docker provider only)
        if len(vm_ip_ports) > 1:
            self.server_port = int(vm_ip_ports[1])
            self.chromium_port = int(vm_ip_ports[2])
            self.vnc_port = int(vm_ip_ports[3])
            self.vlc_port = int(vm_ip_ports[4])
            self.mcp_port = int(vm_ip_ports[5])
        self.controller = PythonController(vm_ip=self.vm_ip, server_port=self.server_port)
        self.setup_controller = SetupController(vm_ip=self.vm_ip, server_port=self.server_port, chromium_port=self.chromium_port, vlc_port=self.vlc_port, cache_dir=self.cache_dir_base, client_password=self.client_password, screen_width=self.screen_width, screen_height=self.screen_height)

    def _revert_to_snapshot(self):
        # Revert to certain snapshot of the virtual machine, and refresh the path to vm and ip of vm
        # due to the fact it could be changed when implemented by cloud services
        path_to_vm = self.provider.revert_to_snapshot(self.path_to_vm, self.snapshot_name)
        if path_to_vm and not path_to_vm == self.path_to_vm:
            # path_to_vm has to be a new path 
            
            self.manager.delete_vm(self.path_to_vm, self.region)
            self.manager.add_vm(path_to_vm, self.region)
            self.manager.occupy_vm(path_to_vm, os.getpid(), self.region)
            self.path_to_vm = path_to_vm

    def _save_state(self, snapshot_name=None):
        # Save the current virtual machine state to a certain snapshot name
        self.provider.save_state(self.path_to_vm, snapshot_name)

    def close(self):
        # Close (release) the virtual machine
        self.provider.stop_emulator(self.path_to_vm)

    def _get_remote_mcp_server_log_tail(self, lines: int = 80) -> str:
        command = (
            "from pathlib import Path; "
            f"path = Path({REMOTE_MCP_SERVER_LOG_PATH!r}); "
            "print(''.join(path.read_text(errors='replace').splitlines(True)[-"
            f"{lines}:]) if path.exists() else '')"
        )
        try:
            return self.controller.execute_python_command(command).get("output", "").strip()
        except Exception as exc:
            logger.info("Failed to read MCP server log tail: %s", exc)
            return ""

    def _is_remote_mcp_server_ready(self, host: str = "127.0.0.1", port: int = 9292) -> bool:
        command = (
            "import socket; "
            "sock = socket.socket(); "
            "sock.settimeout(1.0); "
            f"result = sock.connect_ex(({host!r}, {port})); "
            "sock.close(); "
            "print(result == 0)"
        )
        try:
            output = self.controller.execute_python_command(command).get("output", "").strip()
            return output == "True"
        except Exception as exc:
            logger.info("Failed to probe MCP server readiness: %s", exc)
            return False

    def _wait_for_mcp_server_ready(self, max_attempts: int = 20, sleep_seconds: int = 2) -> bool:
        last_log_tail = None
        for attempt in range(1, max_attempts + 1):
            if self._is_remote_mcp_server_ready():
                log_tail = self._get_remote_mcp_server_log_tail()
                if log_tail:
                    logger.info(
                        "MCP server became ready on attempt %d/%d. Current log tail:\n%s",
                        attempt,
                        max_attempts,
                        log_tail,
                    )
                else:
                    logger.info(
                        "MCP server became ready on attempt %d/%d. Log file is empty.",
                        attempt,
                        max_attempts,
                    )
                return True

            log_tail = self._get_remote_mcp_server_log_tail()
            if log_tail != last_log_tail:
                logger.info(
                    "Waiting for MCP server readiness (%d/%d). Current log tail:\n%s",
                    attempt,
                    max_attempts,
                    log_tail or "<empty log>",
                )
                last_log_tail = log_tail
            else:
                logger.info(
                    "Waiting for MCP server readiness (%d/%d). Log tail unchanged.",
                    attempt,
                    max_attempts,
                )
            time.sleep(sleep_seconds)

        final_log_tail = self._get_remote_mcp_server_log_tail()
        logger.info(
            "MCP server did not become ready after %d attempts. Final log tail:\n%s",
            max_attempts,
            final_log_tail or "<empty log>",
        )
        return False

    def reset(self, task_config: Optional[Dict[str, Any]] = None, seed=None, options=None) -> Dict[str, Any]:
        
        # Reset to certain task in OSWorld
        logger.info("Resetting environment...")
        logger.info("Switching task...")
        logger.info("Setting counters...")
        self._traj_no += 1
        self._step_no = 0
        self.action_history.clear()

        for attempt in range(MAX_RETRIES):
            # Only revert to snapshot if environment has been used (step/setup)
            # This optimization is especially important for cloud providers like AWS
            # where unnecessary snapshot operations are costly and time-consuming
            
            if task_config is not None:
                # Only consider task proxy requirement if proxy is enabled at system level
                task_use_proxy = task_config.get("proxy", False) and self.enable_proxy
                if not self.enable_proxy and task_config.get("proxy", False):
                    logger.info("Task requires proxy but proxy is disabled at system level, ignoring proxy requirement.")
                
                if task_use_proxy != self.current_use_proxy:
                    # keep because get_info_from_website depend on this
                    self.current_use_proxy = task_use_proxy
            
            if self.is_environment_used:
                logger.info("Environment has been used, reverting to snapshot {}...".format(self.snapshot_name))
                self._revert_to_snapshot()
                logger.info("Starting emulator...")
                self._start_emulator()
                logger.info("Emulator started.")
                # Reset the usage flag after reverting
                self.is_environment_used = False
            else:
                logger.info("Environment is clean, skipping snapshot revert (provider: {}).".format(self.provider_name))

            if task_config is not None:
                if task_config.get("proxy", False) and self.enable_proxy:
                    # If using proxy and proxy is enabled, set up the proxy configuration
                    self.setup_controller._proxy_setup(self.client_password)
                self._set_task_info(task_config)
                self.setup_controller.reset_cache_dir(self.cache_dir)
                logger.info("Setting up environment...")
                success = self.setup_controller.setup(self.config, task_config.get("proxy", False) and self.enable_proxy)
                if success:
                    # Mark environment as used when setup is successfully executed
                    if self.config:  # Only mark as used if there were actual setup operations
                        self.is_environment_used = True
                    break
                else:
                    logger.error(
                        "Environment setup failed, retrying (%d/%d)...",
                        attempt + 1,
                        MAX_RETRIES,
                    )
                    time.sleep(5)
            else:
                break
            
        logger.info("Environment setup complete.")

        self.setup_controller._launch_setup('soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault', shell=True)
 
        # Upload tools from autoglm package
        import mm_agents.autoglm
        tool_dir = os.path.join(os.path.dirname(mm_agents.autoglm.__file__), 'tools', 'package')
        for file in os.listdir(tool_dir):
            if os.path.isdir(os.path.join(tool_dir, file)):
                continue
            self.setup_controller._upload_file_setup([{
                "local_path": os.path.join(tool_dir, file),
                "path": os.path.join('/home/user', file)
            }])

        # mcp client
        self.setup_controller._upload_file_setup([{
            "local_path": STATIC_MCP_CLIENT_PATH,
            "path": os.path.join('/home/user', 'osworld_mcp_client.py')
        }])

        # mcp server
        if self.action_space == 'mcp':
            # mcp_server_dir = '/nas/ARPO/OSWorld-AutoGLM/mcp_server'
            mcp_server_dir = STATIC_MCP_DIR 
            
            # 第一步：一次性创建所有需要的目录
            self.setup_controller._launch_setup(
                'mkdir -p /home/user/mcp_server /home/user/mcp_server/tools /home/user/mcp_server/tools/apis /home/user/mcp_server/tools/package',
                shell=True
            )

            # 第二步：批量append上传所有文件
            upload_list = []
            for root, dirs, files in os.walk(mcp_server_dir):
                for file in files:
                    local_path = os.path.join(root, file)
                    rel_path = os.path.relpath(local_path, mcp_server_dir)
                    rel_path_posix = rel_path.replace(os.sep, '/')
                    remote_path = '/home/user/mcp_server/' + rel_path_posix
                    upload_list.append({
                        "local_path": local_path,
                        "path": remote_path
                    })

            # 一次性上传所有文件
            self.setup_controller._upload_file_setup(upload_list)
            self.setup_controller._launch_setup('cd /home/user/mcp_server/ && bash launch_server.sh', shell=True)
            if not self._wait_for_mcp_server_ready():
                logger.error("OSWorld MCP server failed to become ready. See startup logs above.")
            
            # for root, dirs, files in os.walk(mcp_server_dir):
            #     for file in files:
            #         local_path = os.path.join(root, file)
            #         rel_path = os.path.relpath(local_path, mcp_server_dir)
            #         remote_path = os.path.join('/home/user/mcp_server', rel_path)
            #         dir_path = os.path.dirname(remote_path)

            #         self.setup_controller._launch_setup(f'mkdir -p {dir_path}', shell=True)

            #         self.setup_controller._upload_file_setup([{
            #             "local_path": local_path,
            #             "path": remote_path
            #         }])


        # import ipdb
        # ipdb.set_trace()

        time.sleep(5)
        observation = self._get_obs()
        return observation
    

    def get_mcp_tool_list(self, tool_name):
        ENV_SETTING = "import os; os.environ['PATH'] = '/home/user/.nvm/versions/node/v22.18.0/bin:/home/user/.local/bin:' + os.environ['PATH']; "
        # _cmd += ENV_SETTING + "import subprocess; print(subprocess.check_output(['which', 'npx'], text=True).strip())"
        # response = self.controller.execute_python_command(_cmd)['output'].strip()

        command = ENV_SETTING
        command += f"from osworld_mcp_client import *; "
        # print(tool_name)
        # import time
        # time.sleep(1000)
        command += f"OsworldMcpClient.list_tools(tool_name='{tool_name}', shuffle={self.shuffle}, rag={self.rag}, order_param={self.order_param}); "
        # command += f"{class_name}.print_result();"
        tool_list = self.controller.execute_python_command(command)['output'].strip()

        try:
            tool_list = ast.literal_eval(tool_list)
        except Exception as e:
            print(e)
            tool_list = []
            
        # logger.info(f"=========Current Tool List:\n{tool_list}") 
        return tool_list

    def call_mcp_tool(self, name, params):
        ENV_SETTING = "import os; os.environ['PATH'] = '/home/user/.nvm/versions/node/v22.18.0/bin:/home/user/.local/bin:' + os.environ['PATH']; "

        command = ENV_SETTING
        command += f"from osworld_mcp_client import *; "
        command += f"OsworldMcpClient.call_tool(name='{name}', params={str(params)}); "
        # command += f"{class_name}.print_result();"
        controller_result = self.controller.execute_python_command(command)
        # if not isinstance(controller_result, dict):
        #     logger.error(
        #         "MCP python command returned non-dict result for tool %s with params %s: %r",
        #         name,
        #         params,
        #         controller_result,
        #     )
        #     return json.dumps({
        #         "success": False,
        #         "result": None,
        #         "error_message": "Failed to execute MCP command: controller returned no response."
        #     }, ensure_ascii=False)

        response = controller_result.get('output', '')
        if response is None:
            response = ''
        response = str(response).strip()

        if not response:
            return json.dumps({
                "success": False,
                "result": None,
                "error_message": "Empty MCP response."
            }, ensure_ascii=False)

        try:
            parsed_response = ast.literal_eval(response)
        except Exception:
            parsed_response = {
                "success": False,
                "result": None,
                "error_message": f"Failed to parse MCP response: {response}"
            }

        if not (
            isinstance(parsed_response, dict)
            and {"success", "result", "error_message"}.issubset(parsed_response.keys())
        ):
            parsed_response = {
                "success": False,
                "result": None,
                "error_message": f"Unexpected MCP response shape: {parsed_response}"
            }

        return json.dumps(parsed_response, ensure_ascii=False)

    def get_current_apps(self):
        apps_code = r"""import subprocess;
command = "wmctrl -xl";
apps = subprocess.run(command, shell=True, capture_output=True, text=True).stdout.strip().split('\n');
print(apps);"""
        window_code = r"""import subprocess;
command = "wmctrl -a :ACTIVE: -v 2>&1 | grep 'Using window' | awk '{print $3}'";
window_id = subprocess.run(command, shell=True, capture_output=True, text=True).stdout.strip();
print(window_id);"""

        apps = self.controller.execute_python_command(apps_code)['output'].strip()
        apps = ast.literal_eval(apps)
        app_list = {}
        
        for app in apps:
            parts = app.split(maxsplit=4)
            if len(parts) < 4:
                continue
            if parts[1] != '0':
                continue
            window_id = parts[0]
            app_name = '.'.join(parts[2].split('.')[-(math.ceil(parts[2].count('.') / 2)):])
            title = parts[3]
            app_list[window_id] = {
                'app_name': app_name,
                'title': title
            }
        
        cur_id = self.controller.execute_python_command(window_code)['output'].strip()

        return app_list, cur_id

    def maximize_window(self):
        window_state = r"""import subprocess;
command = "xprop -id $(xprop -root _NET_ACTIVE_WINDOW | awk -F' ' '{print $5}') _NET_WM_STATE"
output = subprocess.run(command, shell=True, capture_output=True, text=True).stdout.strip();
print(output);"""
        for _ in range(5):
            try:
                self.setup_controller._launch_setup('wmctrl -r :ACTIVE: -b add,maximized_vert,maximized_horz', shell=True)
                time.sleep(2)
                output = self.controller.execute_python_command(window_state)['output'].strip()
                if '_NET_WM_STATE_FOCUSED' not in output or '_NET_WM_STATE_SKIP_TASKBAR' in output or '_NET_WM_STATE_MODAL' in output or '_NET_WM_STATE_MAXIMIZED' in output: # 没有窗口 or popups or 模态窗口 or 窗口已经最大化
                    return
            except Exception as e:
                logger.error(f"Failed to maximize window: {e}")
                time.sleep(1)

    def _get_obs(self):
        # We provide screenshot, accessibility_tree (optional), terminal (optional), and instruction.
        # can be customized and scaled
        tool_list = {
            "libreoffice_calc": "CalcTools",
            "libreoffice_impress": "ImpressTools",
            "libreoffice_writer": "WriterTools",
            # "code": "CodeTools",
            # "vlc": "VLCTools",
            # "google_chrome": "BrowserTools",
            # "thunderbird": "ThunderbirdTools",
            # "os": "OSTools"
        }
        
        self.maximize_window()

        for i in range(3):
            try:
                app_list, cur_id = self.get_current_apps()
            except Exception as e:
                if i == 2:
                    raise e
                logger.error(f"Failed to get current apps: {e}")
                time.sleep(1)
        
        # import ipdb
        # ipdb.set_trace()
        tool_name = None
        if cur_id in app_list:
            cur_app = app_list[cur_id]['app_name']
            tool_name = cur_app.strip().lower().replace('-', '_')
            if tool_name in tool_list:
                app_info = None
                class_name = tool_list[tool_name]
                command = f"from {tool_name} import *; "
                command += f"{class_name}.env_info(); "
                command += f"{class_name}.print_result();"
                app_info = self.controller.execute_python_command(command)['output'].strip()
                
            else:
                app_info = None
        else:
            cur_app = None
            app_info = None
        
        # tree = self.controller.get_accessibility_tree()
        # screenshot = self.controller.get_screenshot()
        # if screenshot is None:
        #     logger.error("Failed to get screenshot.")
        #     screenshot = b''
        logger.info(f"Current APP: {cur_app} Tool Name: {tool_name}")
        tool_list = []
        if self.action_space == 'mcp':
            tool_list = self.get_mcp_tool_list(tool_name)

        return {
            "screenshot": self.controller.get_screenshot(),
            "accessibility_tree": self.controller.get_accessibility_tree() if self.require_a11y_tree else None,
            "terminal": self.controller.get_terminal_output() if self.require_terminal else None,
            "instruction": self.instruction,
            "apps": app_list,
            "cur_window_id": cur_id,
            "cur_app": cur_app,
            "app_info": app_info,
            "tool_name": tool_name,
            "tool_list": tool_list,
        }

    @property
    def vm_platform(self):
        return self.controller.get_vm_platform()

    @property
    def vm_screen_size(self):
        return self.controller.get_vm_screen_size()

    def _set_task_info(self, task_config: Dict[str, Any]):
        """Set task info (proxy logic is handled in reset method)"""
        self.task_id: str = task_config["id"]
        self.cache_dir: str = os.path.join(self.cache_dir_base, self.task_id)
        os.makedirs(self.cache_dir, exist_ok=True)
        self.instruction = task_config["instruction"]
        self.config = task_config["config"] if "config" in task_config else []
        
        self._set_evaluator_info(task_config)

    def _set_evaluator_info(self, task_config: Dict[str, Any]):
        """Set evaluator information from task config"""
        # evaluator dict
        # func -> metric function string, or list of metric function strings
        # conj -> conjunction of multiple metrics if func is a list with length > 1, "and"/"or"
        # result -> result getter config, or list of result getter configs
        # expected (optional) -> expected getter config, or list of expected getter configs
        # options (optional) -> metric options, or list of metric options
        # if func is a str list, then result, expected (if exists), options (if exists) should also be lists of the same length
        # even if one of the metrics does not need expected or options field, it should be included in the list with None
        self.evaluator = task_config["evaluator"]
        self.metric: Metric = [getattr(metrics, func) for func in self.evaluator["func"]] \
            if isinstance(self.evaluator["func"], list) \
            else getattr(metrics, self.evaluator["func"])
        self.metric_conj: str = self.evaluator.get("conj", "and")  # take conjunction of multiple metrics
        if "result" in self.evaluator and len(self.evaluator["result"]) > 0:
            self.result_getter: Getter = [getattr(getters, "get_{:}".format(res["type"])) for res in
                                          self.evaluator["result"]] \
                if isinstance(self.evaluator["result"], list) \
                else getattr(getters, "get_{:}".format(self.evaluator["result"]["type"]))
        else:
            self.result_getter = [None] * len(self.metric) \
                if isinstance(self.metric, list) \
                else None

        if "expected" in self.evaluator and len(self.evaluator["expected"]) > 0:
            self.expected_getter: Getter = [getattr(getters, "get_{:}".format(exp["type"])) if exp else None for exp in
                                            self.evaluator["expected"]] \
                if isinstance(self.evaluator["expected"], list) \
                else getattr(getters, "get_{:}".format(self.evaluator["expected"]["type"]))
        else:
            self.expected_getter = [None] * len(self.metric) \
                if isinstance(self.metric, list) \
                else None
        self.metric_options: Union[List[Dict[str, Any]], Dict[str, Any]] = [opt if opt else {} for opt in
                                                                            self.evaluator["options"]] \
            if isinstance(self.evaluator.get("options", {}), list) \
            else self.evaluator["options"] \
            if "options" in self.evaluator \
            else [{}] * len(self.metric) \
            if isinstance(self.metric, list) \
            else {}

        assert (not isinstance(self.evaluator["func"], list)
                or (len(self.metric) == len(self.result_getter) == len(self.expected_getter) == len(
                    self.metric_options)))

    def step(self, action, pause=2):
        self._step_no += 1
        self.action_history.append(action)
        
        # Mark environment as used when step is called
        self.is_environment_used = True

        reward = 0  # todo: Define reward calculation for each example
        done = False  # todo: Define episode termination condition for each example
        info = {}
        exe_result = ''
        logger.info(f"Step {self._step_no} in trajectory {self._traj_no} with action: {action}")
        # handle the special actions
        if action in ['WAIT', 'FAIL', 'DONE'] or (type(action) == dict and action['action_type'] in ['WAIT', 'FAIL', 'DONE']):
            if action == 'WAIT':
                time.sleep(pause)
            elif action == 'FAIL':
                done = True
                info = {"fail": True}
            elif action == 'DONE':
                done = True
                info = {"done": True}
 
        if self.action_space == "computer_13":
            # the set of all possible actions defined in the action representation
            self.controller.execute_action(action)
        elif self.action_space in ["pyautogui", "claude_computer_use", "mcp"]:
            if action in ['WAIT', 'FAIL', 'DONE']:
                self.controller.execute_action(action)
            elif "OPEN_CHROME_TAB: " in action:
                # if type(action) == str:
                #     # Fix PyAutoGUI '<' character bug before execution
                #     fixed_command = _fix_pyautogui_less_than_bug(action)
                #     self.controller.execute_python_command(fixed_command)
                # elif type(action) == dict:
                #     # Fix PyAutoGUI '<' character bug before execution
                #     fixed_command = _fix_pyautogui_less_than_bug(action['command'])
                #     self.controller.execute_python_command(fixed_command)
                self.setup_controller._chrome_open_tabs_setup([action.split("OPEN_CHROME_TAB: ")[-1].strip()])
                # print(action.split("OPEN_CHROME_TAB: ")[-1].strip())
            else:
                # the set of all possible python commands insides `pyautogui`
                if type(action) == str:  # agent-s2的框架会走这里
                    # Fix PyAutoGUI '<' character bug before execution
                    fixed_command = _fix_pyautogui_less_than_bug(action)
                    exe_result = self.controller.execute_python_command(fixed_command)
                    try:
                        exe_result = exe_result['output'].strip()
                    except:
                        exe_result = ''
                elif type(action) == dict:
                    if self.action_space == 'mcp':
                        # MCP
                        exe_result = self.call_mcp_tool(
                            name=action['action_type'],
                            params=action['parameters']
                        )
                    else:
                        # Fix PyAutoGUI '<' character bug before execution
                        fixed_command = _fix_pyautogui_less_than_bug(action['command'])
                        self.controller.execute_python_command(fixed_command)

        time.sleep(pause)
        observation = self._get_obs()
        observation['exe_result'] = exe_result

        return observation, reward, done, info

    def evaluate(self):
        """
        Evaluate whether the task is successfully completed.
        """

        postconfig = self.evaluator.get("postconfig", [])
        self.setup_controller.setup(postconfig, self.enable_proxy)
        # Mark environment as used if there were postconfig setup operations
        if postconfig:
            self.is_environment_used = True

        if self.evaluator['func'] == "infeasible":
            if len(self.action_history) > 0:
                last_action = self.action_history[-1]
                if last_action == "FAIL" or (type(last_action) == dict and last_action.get('action_type') == 'FAIL'):
                    return 1
            return 0
        else:
            if len(self.action_history) > 0:
                last_action = self.action_history[-1]
                if last_action == "FAIL" or (type(last_action) == dict and last_action.get('action_type') == 'FAIL'):
                    return 0

        if type(self.metric) == list:
            # Multiple metrics to evaluate whether the task is successfully completed
            results = []
            assert len(self.metric) == len(self.result_getter), "The number of metrics and result getters must be the same"
            if "expected" in self.evaluator:
                assert len(self.metric) == len(self.expected_getter), "The number of metrics and expected getters must be the same"
            for idx, metric in enumerate(self.metric):
                try:
                    config = self.evaluator["result"][idx]
                    result_state = self.result_getter[idx](self, config)
                except FileNotFoundError:
                    logger.error("File not found!")
                    if self.metric_conj == 'and':
                        return 0

                if "expected" in self.evaluator and self.expected_getter and self.evaluator["expected"]:
                    expected_state = self.expected_getter[idx](self, self.evaluator["expected"][idx])
                    metric: int = metric(result_state, expected_state, **self.metric_options[idx])
                else:
                    metric: int = metric(result_state, **self.metric_options[idx])

                if self.metric_conj == 'and' and float(metric) == 0.0:
                    return 0
                elif self.metric_conj == 'or' and float(metric) == 1.0:
                    return 1
                else:
                    results.append(metric)

            return sum(results) / len(results) if self.metric_conj == 'and' else max(results)
        else:
            # Single metric to evaluate whether the task is successfully completed
            try:
                result_state = self.result_getter(self, self.evaluator["result"])
            except FileNotFoundError:
                logger.error("File not found!")
                return 0

            if "expected" in self.evaluator and self.expected_getter and self.evaluator["expected"]:
                expected_state = self.expected_getter(self, self.evaluator["expected"])
                metric: float = self.metric(result_state, expected_state, **self.metric_options)
            else:
                metric: float = self.metric(result_state, **self.metric_options)

        return metric

    def render(self, mode='rgb_array'):
        if mode == 'rgb_array':
            return self.controller.get_screenshot()
        else:
            raise ValueError('Unsupported render mode: {}'.format(mode))
