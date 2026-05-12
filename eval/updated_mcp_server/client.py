import asyncio
from fastmcp import Client
import random
import json

CALL_TOOL_TIMEOUT_SECONDS = 75


class OsworldMcpClient:
    config = {
        "mcpServers": {
            "osworld_mcp": {
                "url": "http://localhost:9292/mcp",
                "transport": "streamable-http"
            },
            "filesystem": {
                "command": "npx",
                "args": [
                    "-y",
                    "@modelcontextprotocol/server-filesystem",
                    "/home/user",
                ]
            },
            "git": {
                "command": "uvx",
                "args": [
                    "mcp-server-git"
                ]
            }
        }
    }

    @staticmethod
    def _build_standard_response(success, result=None, error_message=None):
        return {
            "success": bool(success),
            "result": result,
            "error_message": error_message if not success else None,
        }

    @classmethod
    def _stringify_value(cls, value):
        if value is None:
            return None
        if isinstance(value, str):
            return value
        try:
            return json.dumps(value, ensure_ascii=False)
        except Exception:
            return str(value)

    @classmethod
    def _first_non_empty(cls, *values):
        for value in values:
            if value not in (None, "", [], {}):
                return value
        return None

    @classmethod
    def _looks_like_error_text(cls, value):
        if not isinstance(value, str):
            return False
        normalized = value.strip().lower()
        if not normalized:
            return False
        prefixes = (
            "error",
            "failed",
            "failure",
            "exception",
            "traceback",
            "unhandled exception",
            "unexpected error",
            "❌",
        )
        contains = (
            " not found",
            "could not ",
            "unable to ",
            "invalid ",
            "out of range",
            "already exists",
            "does not exist",
        )
        return normalized.startswith(prefixes) or any(fragment in normalized for fragment in contains)

    @classmethod
    def _normalize_payload(cls, payload, force_error=False, fallback_error=None):
        if isinstance(payload, dict) and {
            "success", "result", "error_message"
        }.issubset(payload.keys()):
            return cls._build_standard_response(
                False if force_error else payload.get("success"),
                payload.get("result"),
                cls._first_non_empty(payload.get("error_message"), fallback_error),
            )

        if isinstance(payload, dict) and "success" in payload:
            success = bool(payload.get("success")) and not force_error
            error_message = cls._first_non_empty(payload.get("error_message"), fallback_error)
            if not error_message and not success:
                for key in ("error", "stderr", "message"):
                    if payload.get(key):
                        error_message = payload.get(key)
                        break
            return cls._build_standard_response(
                success,
                payload.get("result", payload),
                error_message,
            )

        if payload is None:
            return cls._build_standard_response(False, None, fallback_error or "Tool call returned empty payload.")

        if isinstance(payload, str):
            if force_error or cls._looks_like_error_text(payload):
                return cls._build_standard_response(False, None, fallback_error or payload)
            return cls._build_standard_response(True, payload, None)

        if isinstance(payload, dict):
            return cls._build_standard_response(not force_error, None if force_error else payload, fallback_error)

        return cls._build_standard_response(not force_error, None if force_error else {"value": payload}, fallback_error)

    @classmethod
    def _extract_text_from_content_item(cls, item):
        if item is None:
            return None

        text_value = getattr(item, "text", None)
        if text_value is None and isinstance(item, dict):
            text_value = item.get("text")
        if text_value not in (None, ""):
            return text_value

        data_value = getattr(item, "data", None)
        if data_value is None and isinstance(item, dict):
            data_value = item.get("data")
        if data_value not in (None, ""):
            return cls._stringify_value(data_value)

        if isinstance(item, dict):
            return cls._stringify_value(item)

        item_type = getattr(item, "type", None)
        if item_type:
            return cls._stringify_value({"type": item_type})

        return cls._stringify_value(item)

    @classmethod
    def _extract_structured_response(cls, response):
        is_error = bool(
            getattr(response, "is_error", False)
            or getattr(response, "isError", False)
        )

        structured_payload = getattr(response, "structured_content", None)
        if structured_payload is None:
            structured_payload = getattr(response, "structuredContent", None)
        if structured_payload is not None:
            return cls._normalize_payload(
                structured_payload,
                force_error=is_error,
                fallback_error="Tool call returned protocol-level error."
                if is_error else None,
            )

        data_payload = getattr(response, "data", None)
        if data_payload is not None:
            return cls._normalize_payload(
                data_payload,
                force_error=is_error,
                fallback_error="Tool call returned protocol-level error."
                if is_error else None,
            )

        content_payload = getattr(response, "content", None)
        content_text = None
        if content_payload:
            text_parts = []
            for item in content_payload:
                text_value = cls._extract_text_from_content_item(item)
                if text_value:
                    text_parts.append(text_value)

            if text_parts:
                content_text = "\n".join(text_parts)
                try:
                    decoded = json.loads(content_text)
                except Exception:
                    decoded = content_text
                return cls._normalize_payload(
                    decoded,
                    force_error=is_error,
                    fallback_error=content_text if is_error else None,
                )

        if is_error:
            return cls._build_standard_response(
                False,
                None,
                content_text or "Tool call returned protocol-level error without structured content."
            )

        return cls._build_standard_response(
            False,
            None,
            "Tool call returned no structured content."
        )

    # @classmethod
    # def list_tools(cls, tool_name):
    #     async def _list_tools():
    #         client = Client(cls.config)
    #         async with client:
    #             tool_list = await client.list_tools()
    #             tool_list = [{
    #                 "name": tool.name,
    #                 "description": tool.description,
    #                 "parameters": tool.inputSchema
    #             } for tool in tool_list
    #             ]

    #         # filter
    #         result = []
    #         for tool in tool_list:
    #             if (tool_name is not None) and (tool_name in tool['name']):
    #                 result.append(tool)

    #         exclude_apps = [
    #             "libreoffice_calc",
    #             "libreoffice_impress",
    #             "libreoffice_writer",
    #             "code",
    #             "vlc",
    #             "google_chrome",
    #             "thunderbird",
    #         ]
    #         if len(result) == 0:
    #             for tool in tool_list:
    #                 skip = False
    #                 for exclude_app in exclude_apps:
    #                     if exclude_app in tool['name']:
    #                         skip = True
    #                 if not skip:
    #                     result.append(tool)

    #         return result

    #     tool_list = asyncio.run(_list_tools())

    #     print(tool_list)
    #     return tool_list

    @classmethod
    def list_tools(cls, tool_name, shuffle=False, rag=True, order_param=False):
        async def _list_tools():
            client = Client(cls.config)
            async with client:
                tool_list = await client.list_tools()
                tool_list = [{
                    "name": tool.name,
                    "description": tool.description,
                    "parameters": tool.inputSchema
                } for tool in tool_list
                ]

            if rag:
                # filter
                result = []
                for tool in tool_list:
                    if (tool_name is not None) and (tool_name in tool['name']):
                        result.append(tool)

                exclude_apps = [
                    "libreoffice_calc",
                    "libreoffice_impress",
                    "libreoffice_writer",
                    "code",
                    "vlc",
                    "google_chrome",
                    "thunderbird",
                ]
                if len(result) == 0:
                    for tool in tool_list:
                        skip = False
                        for exclude_app in exclude_apps:
                            if exclude_app in tool['name']:
                                skip = True
                        if not skip:
                            result.append(tool)
                return result

            else:
                return tool_list

        tool_list = asyncio.run(_list_tools())

        if shuffle:
            random.shuffle(tool_list)
        if order_param:
            tool_list = sorted(
                tool_list,
                key=lambda tool: (
                    len(tool["parameters"]["properties"]),   # 先按元素个数
                    tool["name"]                             # 再按 name 字典序
                )
            )

        print(tool_list)
        return tool_list

    @classmethod
    def call_tool(cls, name, params=None):
        async def _call_tool():
            client = Client(cls.config)
            try:
                async with client:
                    response = await client.call_tool(
                        name,
                        params or {}
                    )
                return cls._extract_structured_response(response)
            except Exception as e:
                return cls._build_standard_response(
                    False,
                    None,
                    f"{e.__class__.__name__}: {e}"
                )

        response = asyncio.run(_call_tool())

        print(response)
        return response


if __name__ == '__main__':
    OsworldMcpClient.call_tool(
        'VSCodeTools_search_text',
        {
            'text': 'files'
        }
    )
