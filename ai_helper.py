# ai_helper.py
"""LLM helper for explaining roster conflicts and suggesting fixes."""

import ollama


def is_ollama_running():
    """Check if the local Ollama server is available."""
    try:
        ollama.list()
        return True
    except Exception:
        return False


def list_local_models():
    """Return a list of locally available Ollama model names."""
    try:
        response = ollama.list()
        models = response.get("models", [])
        return [m["model"] for m in models] if models else []
    except Exception:
        return []


def suggest_fixes(errors, employees, schedule, preferences, model_name="qwen2.5"):
    """
    Ask a local LLM to explain the current roster conflicts and suggest
    1-2 concrete swaps. Falls back to a rule-based message if the LLM is
    unavailable.
    """
    if not errors:
        return "✅ No conflicts detected. The roster looks good!"

    emp_summary = "\n".join(
        f"- {e['name']}: role={e['role']}, cross-train={e.get('cross_train', [])}"
        for e in employees
    )

    prompt = f"""You are a helpful hotel roster assistant.

The current weekly roster has the following rule violations:
{chr(10).join(f'- {err}' for err in errors)}

Staff available:
{emp_summary}

Explain the most important conflict in plain language, then suggest 1-2 concrete employee swaps or removals that would fix it without creating new violations. Keep your answer short and actionable.
"""

    try:
        response = ollama.chat(
            model=model_name,
            messages=[
                {
                    "role": "system",
                    "content": "You are a hotel roster expert. Be concise and actionable.",
                },
                {"role": "user", "content": prompt},
            ],
        )
        return response["message"]["content"]
    except Exception as e:
        return (
            f"⚠️ The AI helper couldn't reach Ollama ({e}). "
            "Try starting Ollama or check the conflicts manually above."
        )
