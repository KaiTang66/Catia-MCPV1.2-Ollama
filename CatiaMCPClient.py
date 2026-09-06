#2026.6.10
import asyncio
from typing import Optional
from contextlib import AsyncExitStack
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp import types

from anthropic import Anthropic
from anthropic.types import TextBlockParam, CacheControlEphemeralParam
from dotenv import load_dotenv
import sys
import os


load_dotenv() #加载.env的环境变量

#Ollama模型常量
OLLAMA_MODEL = os.getenv("model")
API_KEY = os.getenv("api_key")
BASE_URL= os.getenv("base_url")
MAX_TOOL_TURNS = 1000

        

class CatiaMCPClient:
    def __init__(self):
        # 初始化Session（会话状态）和Client（客户端）
        self.session: Optional[ClientSession] = None
        self.exit_stack = AsyncExitStack()
        self.anthropic = Anthropic(
            base_url=BASE_URL,
            api_key=API_KEY
        )
        self.tools: list
        self.prompts: list
        self.resources: list | None
        self.memories: list[str]=[]

    
    #Server（服务器）连接管理
    async def connect_to_mcpserver(self, server_script_path:str):
        """Connect to an MCP Server

            Args:
            server_script_path: Path to the server script with .py 
        """
        is_python = server_script_path.endswith('.py')
        if not (is_python):
            raise ValueError("Server script must be a .py file")
        
        #和服务器建立stdio通信
        path=Path(server_script_path).resolve()
        server_params = StdioServerParameters(
            command="uv",
            args=["--directory", str(path.parent), "run", path.name],
            env=None
        )

        #使用AsyncExitStack来管理Session周期
        stdio_transport= await self.exit_stack.enter_async_context(stdio_client(server_params))
        self.stdio, self.write = stdio_transport
        self.session = await self.exit_stack.enter_async_context(ClientSession(self.stdio, self.write))
        await self.session.initialize()

        #可用Tool列表
        response= await self.session.list_tools()
        self.tools=response.tools
        print("\nConnected to server with tools:",[tool.name for tool in self.tools])
        #可用Resource列表
        response = await self.session.list_resources()
        self.resources=response.resources
        print("\nConnected to server with resources:",[resource.name for resource in self.resources])
        #可用Prompt列表
        response = await self.session.list_prompts()
        self.prompts=response.prompts
        print("\nConnected to server with prompts:",[prompt.name for prompt in self.prompts])


    #Agent循环
    async def agent_loop(self, query: str)->str|list:
        
        messages=[
            {
                "role":"user",
                "content": query
            }
        ]

        response=await self.session.list_tools()
        #从MCP协议里的通用格式变为Claude API的条用格式
        claude_tools=[
            {
                "name":tool.name,
                "description":tool.description,
                "input_schema":tool.inputSchema
            } for tool in response.tools
        ]

        #初始化Claude API call
        myprompt=""
        for prompt in self.prompts:
            getpromptresult=await self.session.get_prompt(prompt.name)
            for PromptMessage in getpromptresult.messages:
                if isinstance(PromptMessage.content, types.TextContent):
                    myprompt+=PromptMessage.content.text
        system_prompt=f"You are a Catia assistant.\n{myprompt}\n"
        system_prompt+="\n".join(self.memories)
        
        textblockparam=TextBlockParam(
            text=system_prompt,
            type="text",
            cache_control=CacheControlEphemeralParam(type=["ephemeral"],ttl=["1h"])
            )
        
        response=self.anthropic.messages.create(
            model= OLLAMA_MODEL,
            max_tokens= MAX_TOOL_TURNS,
            thinking={"type":"disabled"},
            system=[textblockparam],
            messages=messages,
            tools=claude_tools
        )
        
        #过程回复和Tool调用处理
        
        final_text=[]    
        assistant_message_content=[]
        next_response=True

        while next_response==True:
            
            call_tool_result=True
            
            if response.stop_reason=="end_turn":
                for item in messages:
                    final_text.append(str(item["content"]))
                break
            for i in range(len(response.content)):
                mycontent=response.content[i]
                if mycontent.type == 'text':
                    final_text.append(mycontent.text)
                    assistant_message_content=[mycontent]
                    messages.append(
                        {
                            "role":"assistant",
                            "content":assistant_message_content
                        }
                    )

                elif mycontent.type=='tool_use':
                    tool_name=mycontent.name
                    tool_args=mycontent.input

                    #Tool调用执行
                    result=await self.session.call_tool(tool_name, tool_args)
                    if result.isError==True:
                        call_tool_result = False
                        final_text.append(f"can't use the CatiaMCPTool with {tool_name}")

                    else:    
                        final_text.append(f"use the CatiaMCPTool with {tool_name}")
                    
                    newcontent=[]
                    for item in result.content:
                        if isinstance(item,types.TextContent):
                            final_text.append(item.text)
                            newcontent.append(
                                {
                                    "type":"text",
                                    "text":item.text
                                    }
                            )
                        elif isinstance(item, types.ImageContent):
                            final_text.append(f"the {item.mimeType} with date: {item.data}")
                            newcontent.append(
                                {
                                    "type":"image",
                                    "source":{
                                        "data":item.data,
                                        "media_type":"image/jpeg",
                                        "type":"base64",
                                }
                                }
                            )
                        elif isinstance(item, types.EmbeddedResource):
                            final_text.append(item.resource.text or item.resource.blob)
                            newcontent.append(
                                {
                                    "type":"text",
                                    "text":item.resource.text or item.resource.blob
                                }
                            )
                        
                    assistant_message_content=[mycontent]
                    messages.append(
                        {
                            "role":"assistant",
                            "content":assistant_message_content
                        }
                    )

                    
                    messages.append(
                        {
                            "role":"user",
                            "content":[
                                {
                                    "type":"tool_result",
                                    "tool_use_id":mycontent.id,
                                    "content": newcontent
                                }
                            ]
                        }
                    )
                    
            next_response=False
            if call_tool_result== True:
                #工具调用结果的请求回复
                response=self.anthropic.messages.create(
                    model=OLLAMA_MODEL,
                    max_tokens=MAX_TOOL_TURNS,
                    thinking={"type":"disabled"},
                    system=[textblockparam],
                    messages=messages,
                    tools=claude_tools
                )
                
                
                for i in range(len(response.content)):
                    if response.content[i].type=="tool_use":
                        next_response=True
                        
                
                
            elif call_tool_result == False:
                self.memories+=final_text
                return messages
        for i in range(len(response.content)):
            final_text.append(response.content[i].text)
        self.memories+=final_text
        return "\n".join(final_text)
        

    #对话循环
    async def chat_loop(self):
        """Run an interactive chat loop"""
        print("\nMCP Client Started!")
        print("Type your queries or 'quit' to exit")

        while True:
            try:
                query=input("\nQuery:").strip()
                if query.lower()=='quit':
                    await self.exit_stack.aclose()
                    break
                if query.lower()=='查询记忆':
                    print(self.memories)
                else:
                    self.memories+=query
                    response = await self.agent_loop(query)
                    if isinstance(response, str):
                        print(f"\n{response}")
                    elif isinstance(response, list):
                        print("The Operation form LLM ist not right")
                    else:
                        print(type(response))
            except Exception as e:
                print(f"\nError: {str(e)}")


    #关闭通信
    async def shutdown(self):
        """Clean up resources"""
        await self.exit_stack.aclose()



#主逻辑/Host
async def main():
    if len(sys.argv)<2:
        print("Usage: python CatiaMCPClient.py CatiaMCPServer.py")
        sys.exit(1)#立即结束程序，并返回错误码1
    api_key = API_KEY
    if not api_key:
        print("\nNo ANTHROPIC_API_KEY found. To query these tools with Claude, set your API key:")
        print("  export ANTHROPIC_API_KEY=your-api-key-here")
        return
    client=CatiaMCPClient()
    try:
        await client.connect_to_mcpserver(sys.argv[1])

        await client.chat_loop()
    except Exception as e:
        await client.shutdown()

    

if __name__ == "__main__":
    asyncio.run(main())