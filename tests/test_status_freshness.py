"""Golden test set for check_status_freshness and its supporting helpers.
No real network calls — notary_agent._fetch_page_text is mocked per test.

Deterministic, no-LLM freshness layer: for a card claiming its document is
currently in force ("Статус: действует"), verifies the fetched page doesn't
contain an explicit signal that it's actually been repealed — the master
prompt already defines "Статус"/"Актуальность редакции" as card fields,
this check is the first time either is cross-checked against real content
(see AGENT-BUILDING-METHODOLOGY.md, категория 5, "Верификация
веб-цитирования").
"""
import unittest
from unittest.mock import patch

import notary_agent as na

CARD = (
    "Полное наименование: Конституция Российской Федерации\n"
    "Статус: действует\n"
    "URL2: `https://example.com/constitution`\n"
)


class StatusClaimsInForceTests(unittest.TestCase):
    def test_plain_in_force(self) -> None:
        self.assertTrue(na._status_claims_in_force("действует"))

    def test_in_force_as_of_redaction_date(self) -> None:
        self.assertTrue(na._status_claims_in_force("действует в редакции от 01.01.2026"))

    def test_repealed_variants_are_not_in_force(self) -> None:
        for status in ("утратил силу", "не действует", "утратила силу", "отменён"):
            with self.subTest(status=status):
                self.assertFalse(na._status_claims_in_force(status))

    def test_empty_status_is_not_in_force(self) -> None:
        self.assertFalse(na._status_claims_in_force(""))


class ParseStatusPairsTests(unittest.TestCase):
    def test_extracts_status_and_doc_name(self) -> None:
        pairs = na._parse_status_pairs(CARD)
        self.assertEqual(len(pairs), 1)
        self.assertEqual(pairs[0]["status"], "действует")
        self.assertEqual(pairs[0]["doc_name"], "Конституция Российской Федерации")


class CheckStatusFreshnessTests(unittest.TestCase):
    """Golden cases for check_status_freshness."""

    def test_card_already_claims_repeal_skipped_without_fetch(self) -> None:
        content = CARD.replace("Статус: действует", "Статус: утратил силу")
        with patch.object(na, "_fetch_page_text") as mock_fetch:
            result = na.check_status_freshness(content, 2)
        self.assertEqual(result, [])
        mock_fetch.assert_not_called()

    def test_in_force_claim_confirmed_by_page_no_block(self) -> None:
        with patch.object(na, "_fetch_page_text", return_value=("Конституция Российской Федерации. Глава 1. ...", "ok")):
            result = na.check_status_freshness(CARD, 2)
        self.assertEqual(result, [])

    def test_in_force_claim_contradicted_by_page_hard_block(self) -> None:
        """Catches: a notary cites a repealed act as currently in force."""
        with patch.object(
            na, "_fetch_page_text",
            return_value=(
                "Федеральный закон № 123-ФЗ. Документ утратил силу с 01.01.2020 "
                "в связи с принятием нового закона.",
                "ok",
            ),
        ):
            result = na.check_status_freshness(CARD, 2)
        self.assertEqual(len(result), 1, result)
        msg = result[0]
        self.assertIn("[freshness] БЛОК", msg)
        self.assertIn("утратил силу", msg)
        self.assertIn("КонсультантПлюс", msg)
        self.assertIn("КАРАНТИН", msg)

    def test_unreachable_page_not_this_functions_job(self) -> None:
        with patch.object(na, "_fetch_page_text", return_value=("", "blocked")):
            result = na.check_status_freshness(CARD, 2)
        self.assertEqual(result, [])

    def test_sub_provision_repeal_mention_not_blocked(self) -> None:
        """False-positive guard: a page legitimately describing amendment
        history for ONE provision must not read as whole-document repeal."""
        with patch.object(
            na, "_fetch_page_text",
            return_value=(
                "Конституция Российской Федерации. Глава 1. ... Пункт 5 статьи 22 "
                "утратил силу в связи с принятием Федерального закона № 45-ФЗ.",
                "ok",
            ),
        ):
            result = na.check_status_freshness(CARD, 2)
        self.assertEqual(result, [])

    def test_whole_document_repeal_without_provision_context_still_blocks(self) -> None:
        with patch.object(
            na, "_fetch_page_text",
            return_value=("Конституция Российской Федерации. Настоящий документ утратил силу с 01.01.2020.", "ok"),
        ):
            result = na.check_status_freshness(CARD, 2)
        self.assertEqual(len(result), 1, result)

    def test_part_number_out_of_range_skips_entirely(self) -> None:
        with patch.object(na, "_fetch_page_text") as mock_fetch:
            result = na.check_status_freshness(CARD, 1)
        self.assertEqual(result, [])
        mock_fetch.assert_not_called()


class FindWholeDocumentRepealMarkerTests(unittest.TestCase):
    """Direct tests of the sub-provision filter helper."""

    def test_no_marker_returns_none(self) -> None:
        self.assertIsNone(na._find_whole_document_repeal_marker("документ действует в полном объёме"))

    def test_marker_near_sub_provision_reference_is_skipped(self) -> None:
        text = "пункт 5 статьи 22 утратил силу в связи с изменениями"
        self.assertIsNone(na._find_whole_document_repeal_marker(text))

    def test_marker_without_sub_provision_reference_is_returned(self) -> None:
        text = "настоящий документ утратил силу с 01.01.2020"
        hit = na._find_whole_document_repeal_marker(text)
        self.assertIsNotNone(hit)
        marker, idx = hit
        self.assertEqual(marker, "утратил силу")
        self.assertEqual(text[idx:idx + len(marker)], marker)


if __name__ == "__main__":
    unittest.main()
