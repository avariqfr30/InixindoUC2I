import concurrent.futures
import io
import sys
import threading
import time
import unittest
from pathlib import Path
from unittest import mock


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))


class ChartEngineCacheTests(unittest.TestCase):
    def setUp(self):
        from document_builder import ChartEngine

        ChartEngine._clear_cache()

    def tearDown(self):
        from document_builder import ChartEngine

        ChartEngine._clear_cache()

    def test_identical_calls_render_once_and_return_fresh_streams(self):
        from document_builder import ChartEngine

        payload = b"same-png"
        with mock.patch.object(ChartEngine, "_render_bar_chart", return_value=payload) as render:
            first = ChartEngine.create_bar_chart("Title|Value|A,1", (12, 34, 56))
            second = ChartEngine.create_bar_chart("Title|Value|A,1", [12, 34, 56])

        self.assertEqual(render.call_count, 1)
        self.assertIsNot(first, second)
        self.assertEqual(first.tell(), 0)
        self.assertEqual(second.tell(), 0)
        self.assertEqual(first.read(), payload)
        self.assertEqual(second.read(), payload)
        self.assertEqual(first.tell(), len(payload))
        first.seek(0)
        second.seek(0)
        self.assertEqual(first.tell(), 0)
        self.assertEqual(second.tell(), 0)

    def test_kind_data_and_color_are_separate_cache_keys(self):
        from document_builder import ChartEngine

        with mock.patch.object(ChartEngine, "_render_bar_chart", return_value=b"bar") as bar_render:
            ChartEngine.create_bar_chart("A,1", (1, 2, 3))
            ChartEngine.create_bar_chart("A,2", (1, 2, 3))
            ChartEngine.create_bar_chart("A,1", (3, 2, 1))
        with mock.patch.object(ChartEngine, "_render_line_chart", return_value=b"line") as line_render:
            ChartEngine.create_line_chart("A,1", (1, 2, 3))

        self.assertEqual(bar_render.call_count, 3)
        self.assertEqual(line_render.call_count, 1)

    def test_cache_is_bounded_to_64_entries_with_lru_eviction(self):
        from document_builder import ChartEngine

        with mock.patch.object(
            ChartEngine,
            "_render_bar_chart",
            side_effect=lambda data, color: data.encode("utf-8"),
        ) as render:
            for index in range(64):
                ChartEngine.create_bar_chart(f"A,{index}", (1, 2, 3))
            ChartEngine.create_bar_chart("A,0", (1, 2, 3))
            ChartEngine.create_bar_chart("A,64", (1, 2, 3))
            ChartEngine.create_bar_chart("A,1", (1, 2, 3))

        self.assertEqual(len(ChartEngine._chart_cache), 64)
        self.assertEqual(render.call_count, 66)

    def test_simultaneous_identical_calls_render_once_and_never_overlap_matplotlib(self):
        from document_builder import ChartEngine

        active = 0
        max_active = 0
        call_count = 0
        state_lock = threading.Lock()

        def render(data, color):
            nonlocal active, max_active, call_count
            with state_lock:
                active += 1
                max_active = max(max_active, active)
                call_count += 1
            time.sleep(0.03)
            with state_lock:
                active -= 1
            return b"thread-safe-png"

        with mock.patch.object(ChartEngine, "_render_bar_chart", side_effect=render):
            with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
                streams = list(
                    executor.map(
                        lambda _: ChartEngine.create_bar_chart("A,1", (1, 2, 3)),
                        range(8),
                    )
                )

        self.assertEqual(call_count, 1)
        self.assertEqual(max_active, 1)
        self.assertEqual(len({id(stream) for stream in streams}), 8)
        self.assertTrue(all(stream.tell() == 0 for stream in streams))

    def test_none_and_failed_renders_are_not_cached(self):
        from document_builder import ChartEngine

        with mock.patch.object(ChartEngine, "_render_pie_chart", return_value=None) as render:
            self.assertIsNone(ChartEngine.create_pie_chart("invalid", (1, 2, 3)))
            self.assertIsNone(ChartEngine.create_pie_chart("invalid", (1, 2, 3)))
        self.assertEqual(render.call_count, 2)

        with mock.patch.object(
            ChartEngine,
            "_render_flowchart",
            side_effect=[RuntimeError("render failed"), b"recovered"],
        ) as render:
            with self.assertLogs("document_builder", level="WARNING") as logs:
                self.assertIsNone(ChartEngine.create_flowchart("A->B", (1, 2, 3)))
            recovered = ChartEngine.create_flowchart("A->B", (1, 2, 3))

        self.assertEqual(render.call_count, 2)
        self.assertIn("Gagal membuat flowchart: render failed", "\n".join(logs.output))
        self.assertIsInstance(recovered, io.BytesIO)
        self.assertEqual(recovered.read(), b"recovered")

    def test_all_public_renderers_produce_cached_valid_png_streams(self):
        from document_builder import ChartEngine

        cases = [
            (ChartEngine.create_bar_chart, "Bar|Value|A,1;B,2"),
            (ChartEngine.create_line_chart, "Line|Value|A,1;B,2"),
            (ChartEngine.create_pie_chart, "Pie|A,1;B,2"),
            (ChartEngine.create_flowchart, "Start->Review->Finish"),
        ]

        for renderer, data in cases:
            first = renderer(data, (32, 90, 140))
            second = renderer(data, (32, 90, 140))

            self.assertIsInstance(first, io.BytesIO)
            self.assertIsInstance(second, io.BytesIO)
            self.assertIsNot(first, second)
            self.assertEqual(first.tell(), 0)
            self.assertEqual(second.tell(), 0)
            first_bytes = first.read()
            second_bytes = second.read()
            self.assertTrue(first_bytes.startswith(b"\x89PNG\r\n\x1a\n"))
            self.assertGreater(len(first_bytes), 1_000)
            self.assertEqual(first_bytes, second_bytes)

    def test_all_renderers_close_figures_when_png_save_fails(self):
        import document_builder
        from document_builder import ChartEngine

        cases = [
            (ChartEngine.create_bar_chart, "Bar|Value|A,1;B,2"),
            (ChartEngine.create_line_chart, "Line|Value|A,1;B,2"),
            (ChartEngine.create_pie_chart, "Pie|A,1;B,2"),
            (ChartEngine.create_flowchart, "Start->Finish"),
        ]

        with (
            mock.patch.object(
                document_builder.plt,
                "savefig",
                side_effect=RuntimeError("save failed"),
            ),
            mock.patch.object(
                document_builder.plt,
                "close",
                wraps=document_builder.plt.close,
            ) as close,
        ):
            for renderer, data in cases:
                with self.assertLogs("document_builder", level="WARNING"):
                    self.assertIsNone(renderer(data, (32, 90, 140)))

        self.assertEqual(close.call_count, 4)


if __name__ == "__main__":
    unittest.main()
