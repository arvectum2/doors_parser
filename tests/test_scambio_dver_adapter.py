import unittest

from src.adapters.dver_com import (
    clean_scambio_model_name,
    is_non_door_model_name,
    scambio_collection_queries,
)


class ScambioDverAdapterTests(unittest.TestCase):
    def test_clean_scambio_model_name_removes_sizes_and_sale(self):
        raw = "Дверь ПСТ-4 белый ясень со стеклом Зеркало тонированное размер 200 x 80 Распродажа"
        self.assertEqual(
            clean_scambio_model_name(raw),
            "Дверь ПСТ-4 белый ясень со стеклом Зеркало тонированное",
        )

    def test_clean_scambio_model_name_trims_ellipsis(self):
        raw = "Дверь Смальта-Шарм 12 ... 200х80"
        self.assertEqual(clean_scambio_model_name(raw), "Дверь Смальта-Шарм 12")

    def test_non_door_model_filter(self):
        self.assertTrue(is_non_door_model_name("Коробка прямая МДФ белый 74*33*2070"))
        self.assertTrue(is_non_door_model_name("Профиль в разрезе"))
        self.assertFalse(is_non_door_model_name("Дверь ПСУ-38 Шеллгрей"))

    def test_collection_queries_for_smalt_and_short_lines(self):
        self.assertIn("Смальта-Деко", scambio_collection_queries("Смальта Деко"))
        self.assertIn("Серия ПСЛ", scambio_collection_queries("ПСЛ"))
        self.assertIn("Серия ПСЦ модерн", scambio_collection_queries("ПСЦ Модерн"))
        self.assertIn("Инвизибл", scambio_collection_queries("Инвизибл"))


if __name__ == "__main__":
    unittest.main()
