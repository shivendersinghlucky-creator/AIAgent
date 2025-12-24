"""
🤖 MY FIRST AI AGENT - Test Case Generator
============================================

WHAT IS AN AI AGENT?
--------------------
An AI Agent is like a smart assistant that can:
1. THINK - Understand your request
2. PLAN - Decide what tools to use
3. ACT - Execute actions using tools
4. LEARN - Improve from feedback

Think of it like this:
- Regular Chatbot: You ask → It answers (one-way)
- AI Agent: You ask → It thinks → Uses tools → Takes actions → Achieves goal

REAL-WORLD EXAMPLE:
-------------------
You: "Generate test cases for a login function"
Agent: 
  1. Understands you need test cases
  2. Analyzes what "login" means
  3. Uses its "generate_test_cases" tool
  4. Returns comprehensive test cases
"""

from groq import Groq
import json
from typing import List, Dict
import os

# Initialize Groq client with API key from environment or user input
API_KEY = os.getenv("GROQ_API_KEY", "YOUR_API_KEY_HERE")
client = Groq(api_key=API_KEY)


# ==========================================
# PART 1: TOOLS (Agent's Capabilities)
# ==========================================
# Tools are functions that the agent can use to perform actions
# Think of tools as the agent's "hands" to interact with the world

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "generate_test_cases",
            "description": "Generates comprehensive test cases for a given function or feature. Returns positive, negative, edge, and boundary test cases.",
            "parameters": {
                "type": "object",
                "properties": {
                    "function_name": {
                        "type": "string",
                        "description": "The name of the function or feature to test"
                    },
                    "description": {
                        "type": "string",
                        "description": "Brief description of what the function does"
                    },
                    "parameters": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of function parameters"
                    }
                },
                "required": ["function_name", "description"]
            }
        }
    }
]


def generate_test_cases(function_name: str, description: str, parameters: List[str] = None) -> Dict:
    """
    This is the ACTUAL TOOL that generates test cases.
    When the agent decides to use this tool, this function runs.
    """
    print(f"\n🔧 TOOL ACTIVATED: generate_test_cases")
    print(f"   Function: {function_name}")
    print(f"   Description: {description}")
    
    test_cases = {
        "function": function_name,
        "description": description,
        "test_categories": {
            "positive_tests": [
                {
                    "id": "TC001",
                    "name": f"Valid {function_name} - Happy Path",
                    "description": "Test with valid inputs that should succeed",
                    "input": "Valid test data",
                    "expected": "Success"
                }
            ],
            "negative_tests": [
                {
                    "id": "TC002",
                    "name": f"Invalid {function_name} - Error Handling",
                    "description": "Test with invalid inputs that should fail gracefully",
                    "input": "Invalid test data",
                    "expected": "Error message"
                }
            ],
            "edge_cases": [
                {
                    "id": "TC003",
                    "name": f"Edge Case - {function_name}",
                    "description": "Test boundary conditions and extreme values",
                    "input": "Edge case data",
                    "expected": "Proper handling"
                }
            ],
            "security_tests": [
                {
                    "id": "TC004",
                    "name": f"Security - {function_name}",
                    "description": "Test for SQL injection, XSS, etc.",
                    "input": "Malicious input",
                    "expected": "Blocked/Sanitized"
                }
            ]
        },
        "total_test_cases": 4
    }
    
    return test_cases


# ==========================================
# PART 2: THE AI AGENT CLASS
# ==========================================

class TestCaseAgent:
    """
    This is our AI Agent!
    
    It has:
    - A brain (LLM from Groq)
    - Tools (functions it can use)
    - Memory (conversation history)
    - Logic (decision-making process)
    """
    
    def __init__(self):
        self.client = client
        self.tools = TOOLS
        self.conversation_history = []
        print("\n🤖 Test Case Agent initialized!")
        print("   I can generate test cases for any function or feature.")
        print("   Just tell me what you need!\n")
    
    def think_and_act(self, user_request: str) -> str:
        """
        This is the CORE of the agent!
        
        AGENT LOOP:
        1. Receive user request
        2. Think: What do I need to do?
        3. Decide: Should I use a tool?
        4. Act: Call the tool or respond
        5. Return: Give result to user
        """
        
        print(f"\n{'='*60}")
        print(f"👤 USER REQUEST: {user_request}")
        print(f"{'='*60}\n")
        
        # Add user message to conversation
        self.conversation_history.append({
            "role": "user",
            "content": user_request
        })
        
        # STEP 1: Agent THINKS (asks LLM what to do)
        print("🧠 AGENT THINKING...")
        print("   Analyzing request and deciding which tool to use...\n")
        
        response = self.client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=self.conversation_history,
            tools=self.tools,
            tool_choice="auto",  # Let the agent decide if it needs tools
            temperature=0.7
        )
        
        response_message = response.choices[0].message
        
        # STEP 2: Agent DECIDES - Does it need to use a tool?
        if response_message.tool_calls:
            print("✅ AGENT DECISION: I need to use a tool!")
            
            # Add agent's decision to history
            self.conversation_history.append(response_message)
            
            # STEP 3: Agent ACTS - Execute the tool
            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                function_args = json.loads(tool_call.function.arguments)
                
                print(f"\n🎯 CALLING TOOL: {function_name}")
                print(f"   Arguments: {json.dumps(function_args, indent=2)}\n")
                
                # Execute the actual function
                if function_name == "generate_test_cases":
                    result = generate_test_cases(**function_args)
                    
                    # Add tool result to conversation
                    self.conversation_history.append({
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "name": function_name,
                        "content": json.dumps(result)
                    })
            
            # STEP 4: Agent SYNTHESIZES - Create final response
            print("🔄 AGENT SYNTHESIZING final response...\n")
            
            final_response = self.client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=self.conversation_history,
                temperature=0.7
            )
            
            final_answer = final_response.choices[0].message.content
            self.conversation_history.append({
                "role": "assistant",
                "content": final_answer
            })
            
            return final_answer
        
        else:
            # Agent decided it doesn't need tools
            print("ℹ️  AGENT DECISION: No tool needed, responding directly.\n")
            answer = response_message.content
            self.conversation_history.append({
                "role": "assistant",
                "content": answer
            })
            return answer


# ==========================================
# PART 3: DEMONSTRATION & USAGE
# ==========================================

def main():
    """
    Let's see our AI Agent in action!
    """
    
    print("\n" + "="*70)
    print("🎓 WELCOME TO YOUR FIRST AI AGENT TUTORIAL!")
    print("="*70)
    print("\nYou're about to see an AI Agent work!")
    print("Watch how it:")
    print("  1. Understands your request")
    print("  2. Decides which tool to use")
    print("  3. Executes the tool")
    print("  4. Returns formatted results")
    print("\n" + "="*70 + "\n")
    
    # Create our agent
    agent = TestCaseAgent()
    
    # Ask user for scenario
    scenario = input("Give me the scenario: ")
    
    print(f"\n📝 Generating Test Cases for: {scenario}\n")
    
    result = agent.think_and_act(
        f"Generate test cases for {scenario}"
    )
    
    print("\n" + "="*60)
    print("🎉 AGENT RESPONSE:")
    print("="*60)
    print(result)
    print("\n" + "="*60 + "\n")
    
    # ============================================
    # LEARNING SUMMARY
    # ============================================
    print("\n" + "="*70)
    print("🎓 WHAT YOU JUST LEARNED:")
    print("="*70)
    print("""
1. AI AGENTS have TOOLS (functions they can call)
   - Our agent has the 'generate_test_cases' tool

2. AI AGENTS make DECISIONS
   - They decide WHEN to use tools vs. when to just chat

3. AI AGENTS follow a LOOP:
   - Think → Decide → Act → Respond

4. AI AGENTS maintain MEMORY
   - They remember the conversation context

5. TOOLS are FUNCTIONS
   - Tools are regular Python functions the agent can call
   - The agent decides which tool to use based on user request

WHAT'S HAPPENING UNDER THE HOOD:
---------------------------------
User Request → LLM analyzes → Decides to use tool → 
Calls function → Gets result → Formats response → Returns to user

This is different from a chatbot because:
- Chatbot: Just generates text
- Agent: Thinks, uses tools, takes actions
    """)
    print("="*70 + "\n")
    
    print("\n🎉 Congratulations! You just learned how AI Agents work!")
    print("   Try modifying the code to add your own tools!\n")


if __name__ == "__main__":
    main()
