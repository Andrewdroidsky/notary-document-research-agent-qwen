import unittest

from notary_agent import validate_part_03_applicable_blocks_have_url2


class Part03ValidationTests(unittest.TestCase):
    def test_applicable_block_without_url2_fails(self) -> None:
        text = """I. Базовые отрасли материального права
Статус: НАЙДЕНО.
Документ(ы) / FAIL-SAFE: Основы законодательства РФ о нотариате, ГК РФ.
"""
        issues = validate_part_03_applicable_blocks_have_url2(text)
        self.assertTrue(
            any("Part 3 block `I` is applicable" in issue for issue in issues),
            issues,
        )


if __name__ == "__main__":
    unittest.main()
