"""Golden test set for check_structural_elements_content_support and its
supporting helpers (_html_to_visible_text, _struct_el_variants). No real
network calls — notary_agent._fetch_page_text is mocked per test.

This is the deterministic, no-LLM content-support layer: for a card
claiming a specific структурный элемент (e.g. "Статья 57"), verifies the
fetched page actually contains it — catching a URL2 that resolves to the
right document but not the right article (the "Статья 57 -> table of
contents" incident that motivated this whole check; see
AGENT-BUILDING-METHODOLOGY.md, категория 5, "Верификация веб-цитирования").
"""
import inspect
import unittest
from unittest.mock import patch

import notary_agent as na

CARD = (
    "Полное наименование: Конституция Российской Федерации\n"
    "Структурный элемент: Статья 57\n"
    "URL2: `https://example.com/constitution`\n"
)


class HtmlToVisibleTextTests(unittest.TestCase):
    def test_strips_tags_and_joins_split_phrase(self) -> None:
        html_fragment = "текст ... <b>Статья</b>&nbsp;57. Право на труд ... ещё текст"
        visible = na._html_to_visible_text(html_fragment)
        self.assertIn("статья 57", visible.lower())

    def test_strips_script_content(self) -> None:
        html_fragment = "видимый текст <script>evil()</script> ещё видимый"
        visible = na._html_to_visible_text(html_fragment)
        self.assertNotIn("evil()", visible)


class StructElVariantsTests(unittest.TestCase):
    def test_compound_reference_includes_inner_element(self) -> None:
        variants = na._struct_el_variants("часть 2 статьи 57")
        self.assertIn("статья 57", variants, variants)
        self.assertIn("ст. 57", variants, variants)


class ContentSupportContentSupportTests(unittest.TestCase):
    """Golden cases for check_structural_elements_content_support."""

    def test_present_despite_markup_split_no_block(self) -> None:
        with patch.object(
            na, "_fetch_page_text",
            return_value=(na._html_to_visible_text("... <b>Статья</b>&nbsp;57. Право на труд ..."), "ok"),
        ):
            result = na.check_structural_elements_content_support(CARD, 2)
        self.assertEqual(result, [])

    def test_genuinely_absent_hard_block_with_excerpt_and_karantin(self) -> None:
        """The exact Article-57-resolves-to-table-of-contents case."""
        with patch.object(
            na, "_fetch_page_text",
            return_value=(
                "Конституция РФ. Оглавление. Глава 1. Основы конституционного строя. Глава 2. Права и свободы.",
                "ok",
            ),
        ):
            result = na.check_structural_elements_content_support(CARD, 2)
        self.assertEqual(len(result), 1, result)
        msg = result[0]
        self.assertIn("[content-support] БЛОК", msg)
        self.assertIn("Оглавление", msg, "expected the real page excerpt embedded in the block message")
        self.assertIn("КонсультантПлюс", msg)
        self.assertIn("КАРАНТИН", msg)
        self.assertNotIn("уровень", msg.lower(), "must not offer relabeling the URL2 level as a way out")
        self.assertNotIn("понизьте", msg.lower())

    def test_no_api_key_parameter_in_signature(self) -> None:
        """This check is deliberately deterministic — no OPENAI_API_KEY judge
        call. This project's real workflow never configures that key (manual/
        Codex-driven capture, not the automated run-topic path) — see
        AGENT-BUILDING-METHODOLOGY.md, категория 5, "Верификация
        веб-цитирования" for why a judge-call design was rejected here."""
        sig = inspect.signature(na.check_structural_elements_content_support)
        self.assertEqual(list(sig.parameters), ["content", "part_number"])

    def test_part_number_out_of_range_skips_entirely(self) -> None:
        with patch.object(na, "_fetch_page_text") as mock_fetch:
            result = na.check_structural_elements_content_support(CARD, 1)
        self.assertEqual(result, [])
        mock_fetch.assert_not_called()


class StructuralElementsSoftRegressionTests(unittest.TestCase):
    """check_structural_elements_soft (the cheap stage-1 heuristic, pre-dating
    the hard-block stage 2 above) must keep behaving exactly as before."""

    def test_present_no_warning(self) -> None:
        with patch.object(na, "_fetch_page_text", return_value=("... Статья 57. Право ...", "ok")):
            result = na.check_structural_elements_soft(CARD, 2)
        self.assertEqual(result, [])

    def test_absent_page_ok_warns(self) -> None:
        with patch.object(na, "_fetch_page_text", return_value=("Совсем другой текст без элемента", "ok")):
            result = na.check_structural_elements_soft(CARD, 2)
        self.assertEqual(len(result), 1, result)
        self.assertIn("не найден на странице", result[0])

    def test_blocked_page_silently_skipped(self) -> None:
        with patch.object(na, "_fetch_page_text", return_value=("", "blocked")):
            result = na.check_structural_elements_soft(CARD, 2)
        self.assertEqual(result, [])


if __name__ == "__main__":
    unittest.main()
