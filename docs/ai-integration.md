# AI pipeline integration

VBAlidator was built specifically to sit behind any LLM-VBA code
generator — FormulaBot, Wand.Tools, Ajelix, ExcelBot, V7 Go, plus your
own Claude / GPT pipelines. Treat it as a deterministic compile-safety
gate that turns "hopefully runs" into "verified to compile".

## The contract

`precheck(source, host=…)` returns a `PrecheckResult` with:

| Field | Meaning |
|-------|---------|
| `score` | 0–100 confidence the code compiles cleanly |
| `compile_safe` | True iff zero blocking findings |
| `errors / warnings / info` | Pre-bucketed issue lists |
| `issues` | Flat normalised list (every entry has `rule_id`, `severity`, `category`) |
| `.json()` | Canonical JSON v2 report (versioned schema) |

The result object is **truthy** when `compile_safe`, so:

```python
if precheck(snippet):
    persist(snippet)
else:
    raise GeneratorBugException(...)
```

## Pattern 1 — Claude / Anthropic SDK loop

```python
import anthropic
from vbalidator import precheck

client = anthropic.Anthropic()

def generate_vba(prompt: str, max_attempts: int = 3) -> str:
    history = [{"role": "user", "content": prompt}]
    for attempt in range(max_attempts):
        msg = client.messages.create(
            model="claude-opus-4-7",
            max_tokens=2048,
            messages=history,
        )
        snippet = msg.content[0].text

        result = precheck(snippet, host="excel")
        if result.compile_safe:
            return snippet

        # Feed the analyzer's findings back as the next user turn.
        history.append({"role": "assistant", "content": snippet})
        feedback = "\n".join(
            f"- [{e['rule_id']}] {e['message']}" for e in result.errors
        )
        history.append({"role": "user", "content": (
            f"VBAlidator rejected the previous version with score "
            f"{result.score}/100. Fix every issue and try again:\n{feedback}"
        )})

    raise RuntimeError(f"Gave up after {max_attempts} attempts")
```

## Pattern 2 — OpenAI Agents tool

```python
from openai import OpenAI
from vbalidator import precheck

def vbalidator_tool(snippet: str, host: str = "excel") -> dict:
    """Tool exposed to the model. Returns the JSON v2 report."""
    return precheck(snippet, host=host).json()

# In your agent definition
tools = [{
    "type": "function",
    "name": "vbalidator",
    "description": "Validates a VBA snippet. Returns score + structured issues.",
    "parameters": {
        "type": "object",
        "properties": {
            "snippet": {"type": "string"},
            "host": {"type": "string", "enum": ["excel","word","access","outlook"]},
        },
        "required": ["snippet"]
    }
}]
```

## Pattern 3 — LangChain / LCEL

```python
from langchain_core.runnables import RunnableLambda
from vbalidator import precheck

def gate(snippet: str) -> str:
    res = precheck(snippet, host="excel")
    if not res.compile_safe:
        raise ValueError(f"VBAlidator rejected (score={res.score}): "
                         + ", ".join(e['rule_id'] for e in res.errors))
    return snippet

vba_chain = generator_prompt | llm | RunnableLambda(gate)
```

## Pattern 4 — CI gate for PRs that touch `.bas` / `.cls` / `.frm`

```yaml
name: VBA pre-merge check
on: pull_request
jobs:
  validate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - run: pip install vbalidator
      - run: vbalidator ./vba --host excel --score-threshold 95
```

Pin a high `--score-threshold` to keep the bar above the standard
"compiles" line — it surfaces stylistic issues (`VBA320` Option
Explicit, …) early.

## Strict vs. non-strict

| Mode | Errors | Warnings | Info | `compile_safe` blocked by |
|------|--------|----------|------|---------------------------|
| `strict=True` (default) | counted | counted in score | counted | errors only |
| `strict=False` | counted | ignored in score | ignored | errors only |

`compile_verified` (round-trip) is treated as a hard error in both
modes — VBE itself refused to compile, that's not negotiable.

## Round-trip cross-check

When you can spare a Windows runner with Office installed:

```python
from vbalidator import precheck

result = precheck(snippet, host="excel", roundtrip=True)

# Static + dynamic agree → near-100% confidence
# Static OK, dynamic fails → rare but possible; investigate
# Static fails, dynamic OK → usually a missing host model, file an FP issue
```

The `--roundtrip` CLI flag does the same thing. On non-Windows hosts
the call degrades gracefully to a single `info`-level `VBA_RT000`
notice instead of crashing.

## Token cost / latency

The precheck is local and deterministic — there is no LLM call inside
VBAlidator itself, so you can run it in tight loops without burning
tokens. End-to-end timings on the typical 200-line module:

| Operation | Cold | Warm |
|-----------|------|------|
| `precheck(snippet)` (in-process) | ~50 ms | ~15 ms |
| `vbalidator ./folder` (CLI) | ~150 ms | — |
| `--roundtrip` (Office launch) | ~3 s | ~1 s |
