import unittest

from prepare_competition_data import OUTPUT_FIELDS, transform_rows


class PrepareCompetitionDataTests(unittest.TestCase):
    def test_maps_and_minimizes_fields(self):
        rows = [
            {
                "用户昵称": "用户甲",
                "性别": "男",
                "评论内容": "联系 13812345678，邮箱 test@example.com，详情 https://example.com",
                "点赞数量": "1,200",
                "回复时间": "2025/03/04 10:30:00",
            },
            {
                "用户昵称": "用户乙",
                "性别": "女",
                "评论内容": "联系 13812345678，邮箱 test@example.com，详情 https://example.com",
                "点赞数量": "1200",
                "回复时间": "2025-03-04 10:30:00",
            },
            {"用户昵称": "空评论", "评论内容": "   "},
        ]

        records, stats = transform_rows(rows, "Bilibili")

        self.assertEqual(stats, {"duplicates_removed": 1, "invalid_removed": 1})
        self.assertEqual(list(records[0]), list(OUTPUT_FIELDS))
        self.assertEqual(records[0]["likes"], 1200)
        self.assertEqual(records[0]["timestamp"], "2025-03-04T10:30:00")
        self.assertNotIn("13812345678", records[0]["content"])
        self.assertNotIn("test@example.com", records[0]["content"])
        self.assertNotIn("https://example.com", records[0]["content"])


if __name__ == "__main__":
    unittest.main()
