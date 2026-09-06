#!/usr/bin/env python
"""
check_verification_gates.py

Deterministic Gate report for the content-support and freshness verification
layers in notary_agent.py (check_structural_elements_content_support,
check_status_freshness) — see AGENT-BUILDING-METHODOLOGY.md, категория 5,
"Верификация веб-цитирования", and Verific New/link_citation_verification_tor.md,
раздел 6 "Обязательные acceptance gates" (Gates A-G) in the methodology repo
(C:\\Users\\koper\\Downloads\\GitHub Projects\\5-Day AI Agent Intensive Course).

This is a code-only check, on purpose (same discipline as the methodology
repo's own scripts/check-methodology-structure.sh) — it does not judge
whether the verification logic is *good*, only:
  1. runs the golden test set and requires every test to pass;
  2. mechanically confirms the specific test methods each applicable Gate
     cites actually exist and passed — a Gate row is not "prose that sounds
     right", it is tied to a real, runnable test;
  3. is explicit about which Gates from the ToR are OUT OF SCOPE for this
     script — Gates B/C/D/F cover liveness/provenance, which predate this
     work (Defense 1-3, check_url2_title_audit_at_capture) and are not
     re-verified here. Claiming full Gate coverage when only two of seven
     axes were built would be exactly the kind of overclaim this
     methodology exists to catch.

Run before any commit touching check_structural_elements_content_support or
check_status_freshness.

Usage: python scripts/check_verification_gates.py
Exit code: 0 if the golden set passes and every applicable-Gate test exists
and passed; 1 otherwise.
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

GOLDEN_TEST_MODULES = ["tests.test_content_support", "tests.test_status_freshness"]

# Gate id -> (applicable to this layer?, human description, test ids that
# substantiate it if applicable — must be real, loadable unittest ids).
GATES: list[dict] = [
    {
        "id": "Gate A",
        "applicable": True,
        "description": "Golden-набор классифицирует все кейсы корректно (100%, не порог)",
        "tests": None,  # substantiated by the full run below, not individual ids
    },
    {
        "id": "Gate B",
        "applicable": False,
        "description": "Liveness воспроизводима (тот же URL/состояние -> тот же вердикт)",
        "why_not_applicable": (
            "liveness — реализовано в check_url2_title_audit_at_capture / "
            "verify_and_annotate_url2_titles, не затронуто этой работой"
        ),
    },
    {
        "id": "Gate C",
        "applicable": False,
        "description": "Soft-404 доказан минимум на 3 реальных шаблонах",
        "why_not_applicable": (
            "soft-404 по телу страницы не реализован; текущий liveness-слой "
            "определяет 404/blocked/timeout по HTTP-коду, не по контенту — "
            "честно не покрыто, не входит в объём content-support/freshness"
        ),
    },
    {
        "id": "Gate D",
        "applicable": False,
        "description": "Стоимость/время ограничены, поведение при потолке проверок определено и покрыто тестом",
        "why_not_applicable": (
            "нет LLM-вызовов и внешнего API в этом слое (судья намеренно "
            "убран из дизайна, см. AGENT-BUILDING-METHODOLOGY.md) — бюджетный "
            "потолок неприменим, стоимость = локальный HTTP-фетч"
        ),
    },
    {
        "id": "Gate E",
        "applicable": True,
        "description": "Сбой проверки никогда не даёт молчаливый PASS",
        "tests": [
            "tests.test_content_support.ContentSupportContentSupportTests.test_part_number_out_of_range_skips_entirely",
            "tests.test_status_freshness.CheckStatusFreshnessTests.test_unreachable_page_not_this_functions_job",
        ],
    },
    {
        "id": "Gate F",
        "applicable": False,
        "description": "Анти-подделка лога (независимая перепроверка + детект генератора + анализ временного следа)",
        "why_not_applicable": (
            "provenance — реализовано раньше как Defense 1/2/3 "
            "(check_research_log_url_authenticity, check_tmp_generator_scripts, "
            "check_research_log_timestamp_clustering), не переоценивается этим скриптом"
        ),
    },
    {
        "id": "Gate G",
        "applicable": True,
        "description": "Каждый ручной числовой/эвристический порог имеет golden-кейс по обе стороны границы",
        "tests": [
            "tests.test_content_support.StructElVariantsTests.test_compound_reference_includes_inner_element",
            "tests.test_status_freshness.FindWholeDocumentRepealMarkerTests.test_marker_near_sub_provision_reference_is_skipped",
            "tests.test_status_freshness.FindWholeDocumentRepealMarkerTests.test_marker_without_sub_provision_reference_is_returned",
        ],
    },
]


def _run_golden_suite() -> tuple[bool, dict[str, bool]]:
    """Runs the full golden test set once. Returns (all_passed, {test_id: passed})."""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    for module_name in GOLDEN_TEST_MODULES:
        suite.addTests(loader.loadTestsFromName(module_name))

    # Collect test ids BEFORE running: TestSuite.run() replaces already-run
    # entries with None as a memory optimization, so iterating the suite
    # afterwards silently loses tests instead of raising — collect first.
    all_ids: set[str] = set()

    def _collect(test_suite) -> None:
        for item in test_suite:
            if isinstance(item, unittest.TestSuite):
                _collect(item)
            else:
                all_ids.add(item.id())

    _collect(suite)

    runner = unittest.TextTestRunner(verbosity=0, stream=sys.stderr)
    result = runner.run(suite)

    failed_ids = {t.id() for t, _ in result.failures} | {t.id() for t, _ in result.errors}
    passed_by_id = {test_id: test_id not in failed_ids for test_id in all_ids}
    return result.wasSuccessful(), passed_by_id


def main() -> int:
    all_passed, passed_by_id = _run_golden_suite()

    print(f"\nGolden test set: {len(passed_by_id)} tests, "
          f"{'ALL PASSED' if all_passed else 'FAILURES PRESENT'}\n")

    ok = all_passed
    for gate in GATES:
        if not gate["applicable"]:
            print(f"⚪ {gate['id']}: не применим к этому слою — {gate['why_not_applicable']}")
            continue

        if gate["id"] == "Gate A":
            status = all_passed
            print(f"{'✅' if status else '❌'} {gate['id']}: {gate['description']} "
                  f"({'весь golden-набор зелёный' if status else 'есть провалы — см. вывод выше'})")
            ok = ok and status
            continue

        missing = [t for t in gate["tests"] if t not in passed_by_id]
        failed = [t for t in gate["tests"] if t in passed_by_id and not passed_by_id[t]]
        status = not missing and not failed
        print(f"{'✅' if status else '❌'} {gate['id']}: {gate['description']}")
        for t in gate["tests"]:
            if t in missing:
                print(f"     ❌ MISSING (test not found): {t}")
            elif t in failed:
                print(f"     ❌ FAILED: {t}")
            else:
                print(f"     ✅ {t}")
        ok = ok and status

    print()
    if ok:
        print("Все применимые Gates выполнены. Gates B/C/D/F честно вне объёма этого слоя (см. выше).")
        return 0
    print("Есть невыполненные применимые Gates или провалы в golden-наборе — см. детали выше.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
