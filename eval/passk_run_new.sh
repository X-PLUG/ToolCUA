#!/bin/bash

set -euo pipefail

: "${API_URL:?Set API_URL to your OpenAI-compatible endpoint.}"
: "${API_KEY:=EMPTY}"
: "${DASHSCOPE_API_KEY:?Set DASHSCOPE_API_KEY before running this script.}"
: "${MODELSCOPE_API_KEY:?Set MODELSCOPE_API_KEY before running this script.}"
: "${USER_AGENT_API_KEY:?Set USER_AGENT_API_KEY before running this script.}"
: "${USER_AGENT_BASE_URL:=https://dashscope.aliyuncs.com/compatible-mode/v1}"
: "${USER_AGENT_MODEL:=gpt-4o-2024-05-13}"
: "${PROJECT_DIR:=$(pwd)}"

export API_KEY
export DASHSCOPE_API_KEY
export MODELSCOPE_API_KEY
export USER_AGENT_API_KEY
export USER_AGENT_BASE_URL
export USER_AGENT_MODEL

export ECS_IP=$(curl -s ifconfig.me)

cd "${PROJECT_DIR}"

# >> "${BASE_RESULT_DIR}/0330evaltxt.log" 2>&1 &

START_TRIAL=0
END_TRIAL=3
BASE_RESULT_DIR="result-0418newenv-50steps-eval-5imgs-qwen3vl"

# 定义模型列表
MODELS=(
    "0430_main_v5modify.v1.ab2.add_gui.v2_step19"
)

# 最外层循环：遍历 trial_id
for (( trial_id=START_TRIAL; trial_id<${END_TRIAL}; trial_id++ ))
do
    echo "=========================================================="
    echo "Starting Trial ID: ${trial_id}"
    echo "=========================================================="

    # 外层循环：遍历模型
    for model in "${MODELS[@]}"
    do
        for rep in {1..2}
        do
            # 打印当前正在运行的任务信息
            echo "=========================================================="
            echo "Running Model: ${model}, Trial ID: ${trial_id}, Repetition: ${rep}"

            # 根据 rep 的值来设定 num_envs 的值
            if [ "$rep" -eq 1 ]; then
                num_envs=8
            elif [ "$rep" -eq 2 ]; then
                num_envs=4
            fi
            echo "Setting --num_envs to ${num_envs}"

            # 执行 Python 命令
            python run_multienv_qwen3vl_toolcua_mcp_eval.py \
                --api_url="${API_URL}" \
                --api_key="${API_KEY}" \
                --model="${model}" \
                --test_all_meta_path 'evaluation_examples/test_oswmcp_feasible.json' \
                --num_envs ${num_envs} \
                --action_space mcp \
                --history_n 4 \
                --new_format True \
                --max_steps 50 \
                --result_dir "${BASE_RESULT_DIR}" \
                --trial-id=${trial_id}

            echo "----------------------------------------------------------"
            echo "Finished Model: ${model}, Trial ID: ${trial_id}, Repetition: ${rep}"
            echo "----------------------------------------------------------"
            echo ""

            echo "=========================================================="
            echo "Restarted Docker"
            echo "=========================================================="
            echo ""
            # 停止所有使用 happysixd/osworld-docker 镜像的正在运行的容器
            RUNNING_CONTAINERS=$(docker ps -q --filter "ancestor=happysixd/osworld-docker")
            if [ -n "$RUNNING_CONTAINERS" ]; then
                echo "Stopping containers based on happysixd/osworld-docker..."
                docker stop $RUNNING_CONTAINERS
                echo "Stopped containers: $RUNNING_CONTAINERS"
            else
                echo "No running containers based on happysixd/osworld-docker to stop."
            fi
            # 删除所有使用 happysixd/osworld-docker 镜像的容器（包括已停止的）
            ALL_CONTAINERS=$(docker ps -a -q --filter "ancestor=happysixd/osworld-docker")
            if [ -n "$ALL_CONTAINERS" ]; then
                echo "Removing containers based on happysixd/osworld-docker..."
                docker rm $ALL_CONTAINERS
                echo "Removed containers: $ALL_CONTAINERS"
            else
                echo "No containers based on happysixd/osworld-docker to remove."
            fi

        done

    done

    echo "=========================================================="
    echo "Completed all models for Trial ID: ${trial_id}"
    echo "=========================================================="
    echo ""

    sudo systemctl restart docker

done

echo "All models and trials completed."
