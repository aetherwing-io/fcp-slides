"""Tests for the SlidesAdapter — verb dispatch, undo/redo, lifecycle."""

import os
import tempfile

import pytest

from fcp_core import EventLog, ParsedOp

from fcp_slides.adapter import SlidesAdapter
from fcp_slides.model.snapshot import SnapshotEvent


@pytest.fixture
def adapter():
    return SlidesAdapter()


@pytest.fixture
def model(adapter):
    return adapter.create_empty("Test Deck", {})


@pytest.fixture
def log():
    return EventLog()


class TestLifecycle:
    def test_create_empty(self, adapter):
        model = adapter.create_empty("My Deck", {})
        assert model.title == "My Deck"
        assert model.prs is not None

    def test_serialize_deserialize(self, adapter, model):
        # Add a slide first
        op = ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"})
        result = adapter.dispatch_op(op, model, EventLog())
        assert result.success

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name

        try:
            adapter.serialize(model, path)
            assert os.path.isfile(path)

            model2 = adapter.deserialize(path)
            assert len(model2.prs.slides) == 1
        finally:
            os.unlink(path)

    def test_get_digest(self, adapter, model, log):
        digest = adapter.get_digest(model)
        assert "Slides: 0" in digest

    def test_unknown_verb(self, adapter, model, log):
        op = ParsedOp(verb="nonexistent", positionals=[], params={})
        result = adapter.dispatch_op(op, model, log)
        assert not result.success
        assert "Unknown verb" in result.message


class TestSlideOps:
    def test_slide_add(self, adapter, model, log):
        op = ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success
        assert len(model.prs.slides) == 1

    def test_slide_add_with_layout(self, adapter, model, log):
        op = ParsedOp(verb="slide", positionals=["add"], params={"layout": "title"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success
        assert len(model.prs.slides) == 1

    def test_slide_activate(self, adapter, model, log):
        # Add two slides
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        # Activate first
        op = ParsedOp(verb="slide", positionals=["activate", "1"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success
        assert adapter.index.active_slide == 0

    def test_slide_rename(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        op = ParsedOp(verb="slide", positionals=["rename", "s1", "intro"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_slide_hide_unhide(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(verb="slide", positionals=["hide", "1"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

        op = ParsedOp(verb="slide", positionals=["unhide", "1"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestShapeOps:
    def test_shape_add(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(verb="shape", positionals=["add", "rectangle"], params={"label": "box1"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_textbox_add(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(verb="textbox", positionals=["Hello World"], params={"x": "1in", "y": "1in"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success
        assert "Hello World" in result.message

    def test_shape_add_with_position(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(
            verb="shape",
            positionals=["add", "oval"],
            params={"x": "2in", "y": "3in", "w": "4in", "h": "3in", "label": "circle1"},
        )
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestTextOps:
    def test_placeholder_set(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "title"}), model, log)

        op = ParsedOp(verb="placeholder", positionals=["set", "title", "My Title"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_bullet_add(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(ParsedOp(verb="textbox", positionals=["Items"], params={"label": "list1"}), model, log)

        op = ParsedOp(verb="bullet", positionals=["list1", "First item"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_text_set(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(ParsedOp(verb="textbox", positionals=["Old text"], params={"label": "box1"}), model, log)

        op = ParsedOp(verb="text", positionals=["set", "box1", "New text"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestTableOps:
    def test_table_add(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(verb="table", positionals=["add", "3", "4"], params={"label": "t1"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success
        assert "3x4" in result.message

    def test_table_header(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(ParsedOp(verb="table", positionals=["add", "3", "3"], params={"label": "t1"}), model, log)

        op = ParsedOp(verb="table", positionals=["header", "t1", "Name", "Age", "City"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_table_row(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(ParsedOp(verb="table", positionals=["add", "3", "3"], params={"label": "t1"}), model, log)

        op = ParsedOp(verb="table", positionals=["row", "t1", "1", "Alice", "30", "NYC"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestChartOps:
    def test_chart_add(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(
            verb="chart",
            positionals=["add", "column"],
            params={"label": "c1", "title": "Revenue"},
        )
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_chart_data(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="chart", positionals=["add", "column"], params={"label": "c1"}),
            model, log,
        )

        op = ParsedOp(
            verb="chart",
            positionals=["data", "c1"],
            params={"categories": "Q1,Q2,Q3,Q4", "series": "Revenue", "values": "100,200,150,300"},
        )
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestStyleOps:
    def test_style_fill(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="shape", positionals=["add", "rectangle"], params={"label": "box1"}),
            model, log,
        )

        op = ParsedOp(verb="style", positionals=["box1"], params={"fill": "blue"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_text_style(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="textbox", positionals=["Hello"], params={"label": "tb1"}),
            model, log,
        )

        op = ParsedOp(verb="text-style", positionals=["tb1", "bold"], params={"size": "24", "color": "red"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestLayoutOps:
    def test_align(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="shape", positionals=["add", "rectangle"], params={"label": "a"}),
            model, log,
        )

        op = ParsedOp(verb="align", positionals=["center", "a"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_z_order(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="shape", positionals=["add", "rectangle"], params={"label": "a"}),
            model, log,
        )

        op = ParsedOp(verb="z-order", positionals=["front", "a"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestNotesOps:
    def test_notes_set(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)

        op = ParsedOp(verb="notes", positionals=["set", "Speaker notes here"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success

    def test_deck_widescreen(self, adapter, model, log):
        op = ParsedOp(verb="deck", positionals=["widescreen"])
        result = adapter.dispatch_op(op, model, log)
        assert result.success


class TestUndoRedo:
    def test_undo_slide_add(self, adapter, model, log):
        # Add a slide
        op = ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"})
        result = adapter.dispatch_op(op, model, log)
        assert result.success
        assert len(model.prs.slides) == 1

        # Undo
        events = log.undo()
        assert len(events) == 1
        adapter.reverse_event(events[0], model)
        assert len(model.prs.slides) == 0

    def test_redo_slide_add(self, adapter, model, log):
        # Add a slide
        op = ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"})
        adapter.dispatch_op(op, model, log)
        assert len(model.prs.slides) == 1

        # Undo
        events = log.undo()
        adapter.reverse_event(events[0], model)
        assert len(model.prs.slides) == 0

        # Redo
        redo_events = log.redo()
        adapter.replay_event(redo_events[0], model)
        assert len(model.prs.slides) == 1


class TestQueries:
    def test_query_plan(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        result = adapter.dispatch_query("plan", model)
        assert "Slide 1" in result

    def test_query_status(self, adapter, model, log):
        result = adapter.dispatch_query("status", model)
        assert "Slides:" in result

    def test_query_list_slides(self, adapter, model, log):
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        result = adapter.dispatch_query("list slides", model)
        assert "Blank" in result or "1." in result

    def test_query_unknown(self, adapter, model, log):
        result = adapter.dispatch_query("nonexistent", model)
        assert "Unknown query" in result


class TestIntegration:
    """End-to-end test simulating a real deck workflow."""

    def test_pitch_deck(self, adapter, log):
        model = adapter.create_empty("Q4 Report", {})

        # Slide 1: Title slide
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "title"}), model, log)
        adapter.dispatch_op(ParsedOp(verb="placeholder", positionals=["set", "title", "Q4 Revenue Report"]), model, log)
        adapter.dispatch_op(ParsedOp(verb="placeholder", positionals=["set", "subtitle", "Prepared for Board"]), model, log)

        # Slide 2: Chart
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="chart", positionals=["add", "column"], params={"label": "revenue", "title": "Revenue by Region"}),
            model, log,
        )
        adapter.dispatch_op(
            ParsedOp(
                verb="chart",
                positionals=["data", "revenue"],
                params={"categories": "North,South,East,West", "series": "Q4", "values": "1.25M,980K,1.1M,870K"},
            ),
            model, log,
        )

        # Slide 3: Table
        adapter.dispatch_op(ParsedOp(verb="slide", positionals=["add"], params={"layout": "blank"}), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="table", positionals=["add", "4", "4"], params={"label": "metrics"}),
            model, log,
        )
        adapter.dispatch_op(
            ParsedOp(verb="table", positionals=["header", "metrics", "Metric", "Q3", "Q4", "Change"]),
            model, log,
        )
        adapter.dispatch_op(
            ParsedOp(verb="table", positionals=["row", "metrics", "1", "Revenue", "$1.3M", "$1.8M", "+38%"]),
            model, log,
        )

        # Verify
        assert len(model.prs.slides) == 3
        digest = adapter.get_digest(model)
        assert "Slides: 3" in digest

        # Save and reload
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name

        try:
            adapter.serialize(model, path)
            model2 = adapter.deserialize(path)
            assert len(model2.prs.slides) == 3
        finally:
            os.unlink(path)
