import os
import time
import re
from io import BytesIO
from typing import Dict, List
from textwrap import dedent
from functools import lru_cache

import streamlit as st
from langgraph.graph import StateGraph, END
from langchain_core.messages import HumanMessage, AIMessage
from langchain_core.runnables import RunnableLambda
from langchain_openai import ChatOpenAI
from tavily import TavilyClient

# Handle DOCX import gracefully
try:
    from docx import Document
    from docx.shared import Inches
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ===================
# Agent Implementation
# ===================

class MyLangGraphAgent:
    def __init__(self, max_retries: int = 3):
        self.max_retries = max_retries
        self.tavily = TavilyClient(api_key=st.secrets["TAVILY_API_KEY"])
        self.llm = self._setup_llm()
        self.graph = self._setup_graph()

    def _setup_llm(self) -> ChatOpenAI:
        return ChatOpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=st.secrets["OPENAI_API_KEY"],
            model="openai/gpt-4o",
            temperature=0.3,
            max_tokens=2000,
            timeout=30,
            max_retries=3
        )

    @lru_cache(maxsize=100)
    def _search_web(self, job: str) -> str:
        """Cached search function with retry logic"""
        query = (
            f'site:forbes.com OR site:techcrunch.com OR site:mit.edu OR site:gartner.com '
            f'("AI" OR "artificial intelligence") AND ("{job}" OR "{job} industry") '
            f'AND ("use case" OR "application" OR "implementation" OR "case study") '
            f'after:2023-01-01'
        )
        
        for attempt in range(self.max_retries):
            try:
                result = self.tavily.search(
                    query=query,
                    search_depth="advanced",
                    max_results=7,
                    include_raw_content=True
                )
                if not result["results"]:
                    return "No results found."
                return "\n\n".join([
                    f"{r['title']}:\n{r['content']}\nSource: {r['url']}" 
                    for r in result["results"][:5]
                )
            except Exception as e:
                if attempt == self.max_retries - 1:
                    return f"Search failed after {self.max_retries} attempts: {str(e)}"
                time.sleep(2 ** attempt)

    def _format_prompt(self, job: str, search_results: str) -> List[Dict[str, str]]:
        system_prompt = dedent(f"""
        You are an expert AI research analyst specializing in industry applications of AI.
        Analyze the following information about AI in {job} and extract specific examples.

        ### Required Output Format ###
        You MUST respond with ONLY this Markdown table structure:

        ```markdown
        | Task/Function | AI Technology | Implementation Details | Impact | Source |
        |---------------|---------------|------------------------|--------|--------|
        [5-8 specific examples]
        ```

        Focus on:
        - Concrete implementations (specific tools/technologies)
        - Measurable impact (include numbers when possible)
        - Verified information (must match the sources provided)
        """)
        return [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": search_results}
        ]

    def _run_agent(self, state: Dict) -> Dict:
        """Wrapper with retry logic for LLM calls"""
        for attempt in range(self.max_retries):
            try:
                messages = state["messages"]
                response = self.llm.invoke(messages)
                return {"messages": messages + [AIMessage(content=response.content)]}
            except Exception as e:
                if attempt == self.max_retries - 1:
                    raise
                time.sleep(1 * (attempt + 1))

    def _setup_graph(self):
        builder = StateGraph(dict)
        builder.add_node("agent_node", RunnableLambda(self._run_agent))
        builder.set_entry_point("agent_node")
        builder.add_edge("agent_node", END)
        return builder.compile()

    def query(self, job: str) -> Dict:
        """Main query method with comprehensive error handling"""
        try:
            if not job or not isinstance(job, str):
                raise ValueError("Job title must be a non-empty string")
                
            search_results = self._search_web(job)
            if "failed" in search_results.lower():
                return {"error": search_results}
                
            prompt = self._format_prompt(job, search_results)
            messages = [HumanMessage(**m) for m in prompt]
            
            return self.graph.invoke(
                {"messages": messages},
                config={"recursion_limit": 50}
            )
        except Exception as e:
            return {
                "error": f"Query failed: {str(e)}",
                "suggestion": "Please try again or check your inputs"
            }

# ===================
# Streamlit App
# ===================

def main():
    st.set_page_config(
        page_title="AI Industry Assistant",
        page_icon="🚀",
        layout="centered"
    )
    
    st.title("🚀 AI in Industry Assistant")
    st.markdown("Discover how AI is transforming specific professions or industries.")
    
    job = st.text_input(
        "Enter a profession or industry:",
        placeholder="e.g., Digital Marketing, Healthcare, Finance",
        help="Be as specific as possible for better results"
    )
    
    if st.button("Analyze") and job:
        with st.spinner(f"Researching AI applications in {job}..."):
            try:
                agent = MyLangGraphAgent()
                result = agent.query(job)
                
                if "error" in result:
                    st.error(result["error"])
                    if "suggestion" in result:
                        st.info(result["suggestion"])
                else:
                    markdown_output = result["messages"][-1].content
                    st.subheader(f"AI Applications in {job}")
                    st.markdown(markdown_output)
                    
                    if DOCX_AVAILABLE:
                        try:
                            from io import BytesIO
                            docx_file, filename = convert_to_docx(markdown_output)
                            st.download_button(
                                label="📄 Download Report",
                                data=docx_file,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        except Exception as e:
                            st.warning(f"DOCX export failed: {str(e)}")
                    else:
                        st.warning("DOCX export unavailable - showing Markdown only")
                        
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
