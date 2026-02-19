"""Tests for watchdog event handlers and WatchDog service lifecycle.

All COM/comtypes dependencies are mocked at the module level so these tests
run without a live Windows UI Automation environment.
"""

import sys
import weakref
from types import ModuleType
from unittest.mock import MagicMock, patch

# ---------------------------------------------------------------------------
# Module-level COM stubs
# These must be in place before any watchdog submodule is imported in a test.
# ---------------------------------------------------------------------------


def _make_comtypes_stub() -> ModuleType:
    """Return a lightweight comtypes stub sufficient for watchdog imports."""
    stub = ModuleType("comtypes")
    stub.CoInitialize = MagicMock()
    stub.CoUninitialize = MagicMock()

    # comtypes.COMObject -- base class used by event handler classes
    class _COMObject:
        def __init__(self):
            pass

    stub.COMObject = _COMObject

    # comtypes.client sub-module
    client_stub = ModuleType("comtypes.client")
    client_stub.PumpEvents = MagicMock()
    stub.client = client_stub
    sys.modules["comtypes.client"] = client_stub

    return stub


def _make_uia_stub() -> tuple[ModuleType, ModuleType]:
    """Return (windows_mcp.uia.core stub, windows_mcp.uia.enums stub)."""
    # Build a fake UIAutomationCore with the interfaces needed by event_handlers
    uia_core_ns = MagicMock()
    uia_core_ns.IUIAutomationFocusChangedEventHandler = object
    uia_core_ns.IUIAutomationStructureChangedEventHandler = object
    uia_core_ns.IUIAutomationPropertyChangedEventHandler = object

    fake_uia_client = MagicMock()
    fake_uia_client.UIAutomationCore = uia_core_ns
    fake_uia_client.IUIAutomation = MagicMock()

    core_stub = ModuleType("windows_mcp.uia.core")
    core_stub._AutomationClient = MagicMock()
    core_stub._AutomationClient.instance.return_value = fake_uia_client

    enums_stub = ModuleType("windows_mcp.uia.enums")
    enums_stub.TreeScope = MagicMock()
    enums_stub.TreeScope.TreeScope_Subtree = 4

    return core_stub, enums_stub, fake_uia_client


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_module_patches(comtypes_stub, core_stub, enums_stub):
    """Return a dict suitable for patch.dict(sys.modules, ...)."""
    return {
        "comtypes": comtypes_stub,
        "comtypes.client": comtypes_stub.client,
        "windows_mcp.uia": ModuleType("windows_mcp.uia"),
        "windows_mcp.uia.core": core_stub,
        "windows_mcp.uia.enums": enums_stub,
    }


# ---------------------------------------------------------------------------
# Event handler tests
# ---------------------------------------------------------------------------


class TestFocusChangedEventHandler:
    """Unit tests for FocusChangedEventHandler (lines 15-29 of event_handlers.py)."""

    def _import_handler(self, patches):
        # Remove cached module so it re-imports with the patched stubs
        sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
        with patch.dict(sys.modules, patches):
            from windows_mcp.watchdog.event_handlers import FocusChangedEventHandler

            return FocusChangedEventHandler

    def test_calls_focus_callback_when_parent_alive(self):
        """HandleFocusChangedEvent invokes _focus_callback on the live parent."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        FocusChangedEventHandler = self._import_handler(patches)

        callback = MagicMock()
        parent = MagicMock()
        parent._focus_callback = callback

        handler = FocusChangedEventHandler.__new__(FocusChangedEventHandler)

        handler._parent = weakref.ref(parent)

        sender = MagicMock()
        result = handler.HandleFocusChangedEvent(sender)

        callback.assert_called_once_with(sender)
        assert result == 0  # S_OK

    def test_skips_callback_when_parent_dead(self):
        """HandleFocusChangedEvent is a no-op when the weak-ref parent is gone."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        FocusChangedEventHandler = self._import_handler(patches)

        handler = FocusChangedEventHandler.__new__(FocusChangedEventHandler)
        # Simulate a dead weak reference
        handler._parent = lambda: None

        result = handler.HandleFocusChangedEvent(MagicMock())
        assert result == 0  # S_OK -- no exception raised

    def test_skips_callback_when_focus_callback_is_none(self):
        """HandleFocusChangedEvent is a no-op when _focus_callback is None."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        FocusChangedEventHandler = self._import_handler(patches)

        parent = MagicMock()
        parent._focus_callback = None

        handler = FocusChangedEventHandler.__new__(FocusChangedEventHandler)

        handler._parent = weakref.ref(parent)

        result = handler.HandleFocusChangedEvent(MagicMock())
        assert result == 0

    def test_callback_exception_is_swallowed(self):
        """Exceptions inside the callback must not propagate -- returns S_OK."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        FocusChangedEventHandler = self._import_handler(patches)

        callback = MagicMock(side_effect=RuntimeError("boom"))
        parent = MagicMock()
        parent._focus_callback = callback

        handler = FocusChangedEventHandler.__new__(FocusChangedEventHandler)

        handler._parent = weakref.ref(parent)

        result = handler.HandleFocusChangedEvent(MagicMock())
        assert result == 0  # Must not raise


class TestStructureChangedEventHandler:
    """Unit tests for StructureChangedEventHandler (lines 32-46 of event_handlers.py)."""

    def _import_handler(self, patches):
        sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
        with patch.dict(sys.modules, patches):
            from windows_mcp.watchdog.event_handlers import StructureChangedEventHandler

            return StructureChangedEventHandler

    def test_calls_structure_callback(self):
        """HandleStructureChangedEvent passes all three arguments to the callback."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        StructureChangedEventHandler = self._import_handler(patches)

        callback = MagicMock()
        parent = MagicMock()
        parent._structure_callback = callback

        handler = StructureChangedEventHandler.__new__(StructureChangedEventHandler)

        handler._parent = weakref.ref(parent)

        sender = MagicMock()
        change_type = 1
        runtime_id = [42]
        result = handler.HandleStructureChangedEvent(sender, change_type, runtime_id)

        callback.assert_called_once_with(sender, change_type, runtime_id)
        assert result == 0

    def test_callback_exception_is_swallowed(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        StructureChangedEventHandler = self._import_handler(patches)

        callback = MagicMock(side_effect=ValueError("bad structure"))
        parent = MagicMock()
        parent._structure_callback = callback

        handler = StructureChangedEventHandler.__new__(StructureChangedEventHandler)

        handler._parent = weakref.ref(parent)

        result = handler.HandleStructureChangedEvent(MagicMock(), 0, [])
        assert result == 0

    def test_skips_callback_when_parent_dead(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        StructureChangedEventHandler = self._import_handler(patches)

        handler = StructureChangedEventHandler.__new__(StructureChangedEventHandler)
        handler._parent = lambda: None

        result = handler.HandleStructureChangedEvent(MagicMock(), 0, [])
        assert result == 0


class TestPropertyChangedEventHandler:
    """Unit tests for PropertyChangedEventHandler (lines 49-63 of event_handlers.py)."""

    def _import_handler(self, patches):
        sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
        with patch.dict(sys.modules, patches):
            from windows_mcp.watchdog.event_handlers import PropertyChangedEventHandler

            return PropertyChangedEventHandler

    def test_calls_property_callback(self):
        """HandlePropertyChangedEvent forwards sender, propertyId, newValue to callback."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        PropertyChangedEventHandler = self._import_handler(patches)

        callback = MagicMock()
        parent = MagicMock()
        parent._property_callback = callback

        handler = PropertyChangedEventHandler.__new__(PropertyChangedEventHandler)

        handler._parent = weakref.ref(parent)

        sender = MagicMock()
        prop_id = 30005  # UIA_NamePropertyId
        new_value = "new name"
        result = handler.HandlePropertyChangedEvent(sender, prop_id, new_value)

        callback.assert_called_once_with(sender, prop_id, new_value)
        assert result == 0

    def test_callback_exception_is_swallowed(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        PropertyChangedEventHandler = self._import_handler(patches)

        callback = MagicMock(side_effect=TypeError("bad prop"))
        parent = MagicMock()
        parent._property_callback = callback

        handler = PropertyChangedEventHandler.__new__(PropertyChangedEventHandler)

        handler._parent = weakref.ref(parent)

        result = handler.HandlePropertyChangedEvent(MagicMock(), 30005, "val")
        assert result == 0

    def test_skips_callback_when_property_callback_is_none(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        PropertyChangedEventHandler = self._import_handler(patches)

        parent = MagicMock()
        parent._property_callback = None

        handler = PropertyChangedEventHandler.__new__(PropertyChangedEventHandler)

        handler._parent = weakref.ref(parent)

        result = handler.HandlePropertyChangedEvent(MagicMock(), 30005, "val")
        assert result == 0


# ---------------------------------------------------------------------------
# Event handler __init__ coverage (lines 19-20, 36-37, 53-54)
# ---------------------------------------------------------------------------


class TestEventHandlerInit:
    """Cover __init__ methods on all three event handler classes.

    Previous tests used __new__ to bypass __init__, leaving the
    constructor lines (weakref + super().__init__()) uncovered.
    """

    def _import_all(self, patches):
        sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
        with patch.dict(sys.modules, patches):
            from windows_mcp.watchdog.event_handlers import (
                FocusChangedEventHandler,
                PropertyChangedEventHandler,
                StructureChangedEventHandler,
            )

            return (
                FocusChangedEventHandler,
                StructureChangedEventHandler,
                PropertyChangedEventHandler,
            )

    def test_focus_handler_init_sets_weakref(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        Focus, _, _ = self._import_all(patches)
        parent = MagicMock()
        handler = Focus(parent)
        assert handler._parent() is parent

    def test_structure_handler_init_sets_weakref(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        _, Structure, _ = self._import_all(patches)
        parent = MagicMock()
        handler = Structure(parent)
        assert handler._parent() is parent

    def test_property_handler_init_sets_weakref(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        _, _, Property = self._import_all(patches)
        parent = MagicMock()
        handler = Property(parent)
        assert handler._parent() is parent


# ---------------------------------------------------------------------------
# WatchDog service tests
# ---------------------------------------------------------------------------


def _fresh_watchdog(patches):
    """Import and return a new WatchDog instance with patched COM dependencies."""
    sys.modules.pop("windows_mcp.watchdog.service", None)
    sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
    sys.modules.pop("windows_mcp.watchdog", None)
    with patch.dict(sys.modules, patches):
        from windows_mcp.watchdog.service import WatchDog

        return WatchDog()


class TestWatchDogInit:
    """Test WatchDog initialisation state."""

    def test_initial_state(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, fake_uia_client = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)

        assert not wd.is_running.is_set()
        assert wd.thread is None
        assert wd._focus_callback is None
        assert wd._structure_callback is None
        assert wd._property_callback is None
        assert wd._focus_handler is None
        assert wd._structure_handler is None
        assert wd._property_handler is None


class TestWatchDogStartStop:
    """Test WatchDog start / stop lifecycle."""

    def test_start_sets_running_flag(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, fake_uia_client = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        with patch.dict(sys.modules, patches):
            sys.modules.pop("windows_mcp.watchdog.service", None)
            sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
            from windows_mcp.watchdog.service import WatchDog

            wd = WatchDog()
            # Patch _run so the thread exits immediately
            wd._run = MagicMock()
            wd.start()

        assert wd.is_running.is_set()
        assert wd.thread is not None

    def test_start_is_idempotent(self):
        """Calling start() twice must not create a second thread."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        with patch.dict(sys.modules, patches):
            sys.modules.pop("windows_mcp.watchdog.service", None)
            sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
            from windows_mcp.watchdog.service import WatchDog

            wd = WatchDog()
            wd._run = MagicMock()
            wd.start()
            first_thread = wd.thread
            wd.start()  # second call should be a no-op
            assert wd.thread is first_thread

    def test_stop_clears_running_flag(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        with patch.dict(sys.modules, patches):
            sys.modules.pop("windows_mcp.watchdog.service", None)
            sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
            from windows_mcp.watchdog.service import WatchDog

            wd = WatchDog()
            wd._run = MagicMock()
            wd.start()
            wd.stop()

        assert not wd.is_running.is_set()

    def test_stop_when_not_running_is_safe(self):
        """stop() on a never-started WatchDog must not raise."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        with patch.dict(sys.modules, patches):
            sys.modules.pop("windows_mcp.watchdog.service", None)
            sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
            from windows_mcp.watchdog.service import WatchDog

            wd = WatchDog()
            wd.stop()  # Should not raise

    def test_context_manager_starts_and_stops(self):
        """Using WatchDog as a context manager calls start() then stop()."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        with patch.dict(sys.modules, patches):
            sys.modules.pop("windows_mcp.watchdog.service", None)
            sys.modules.pop("windows_mcp.watchdog.event_handlers", None)
            from windows_mcp.watchdog.service import WatchDog

            wd = WatchDog()
            wd._run = MagicMock()

            with wd:
                assert wd.is_running.is_set()

            assert not wd.is_running.is_set()


class TestWatchDogCallbackSetters:
    """Test the public callback setter methods."""

    def test_set_focus_callback(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)
        cb = MagicMock()
        wd.set_focus_callback(cb)
        assert wd._focus_callback is cb

    def test_clear_focus_callback(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)
        wd.set_focus_callback(MagicMock())
        wd.set_focus_callback(None)
        assert wd._focus_callback is None

    def test_set_structure_callback_with_element(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)
        cb = MagicMock()
        element = MagicMock()
        wd.set_structure_callback(cb, element=element)
        assert wd._structure_callback is cb
        assert wd._structure_element is element

    def test_set_structure_callback_without_element(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)
        cb = MagicMock()
        wd.set_structure_callback(cb)
        assert wd._structure_callback is cb
        assert wd._structure_element is None

    def test_set_property_callback_with_element_and_ids(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)
        cb = MagicMock()
        element = MagicMock()
        ids = [30005, 30045]
        wd.set_property_callback(cb, element=element, property_ids=ids)
        assert wd._property_callback is cb
        assert wd._property_element is element
        assert wd._property_ids == ids

    def test_clear_property_callback(self):
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, _ = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        wd = _fresh_watchdog(patches)
        wd.set_property_callback(MagicMock(), property_ids=[30005])
        wd.set_property_callback(None)
        assert wd._property_callback is None
        assert wd._property_ids is None


class TestWatchDogRunLoop:
    """Test the _run() main loop logic by patching attributes on the loaded service module.

    We use setattr-based patching on the already-imported module object to avoid
    re-triggering the package __init__.py import chain (which calls live COM APIs).
    """

    def _setup(self):
        """Return (wd, fake_uia_client, service_mod, comtypes_stub) for run-loop tests."""
        comtypes_stub = _make_comtypes_stub()
        core_stub, enums_stub, fake_uia_client = _make_uia_stub()
        patches = _build_module_patches(comtypes_stub, core_stub, enums_stub)

        # Clear cached watchdog modules so a fresh import uses our stubs
        for key in list(sys.modules):
            if "windows_mcp.watchdog" in key:
                del sys.modules[key]

        with patch.dict(sys.modules, patches):
            import windows_mcp.watchdog.service as service_mod

            wd = service_mod.WatchDog()

        return wd, fake_uia_client, service_mod, comtypes_stub

    def _run_one_iteration(self, wd, service_mod, comtypes_stub):
        """Patch handler constructors and comtypes on the module, then run one loop iteration."""
        mock_comtypes = MagicMock()
        mock_comtypes.CoInitialize = MagicMock()
        mock_comtypes.CoUninitialize = MagicMock()
        mock_comtypes.client = comtypes_stub.client

        def pump_and_stop(d):
            wd.is_running.clear()

        comtypes_stub.client.PumpEvents.side_effect = pump_and_stop

        orig_comtypes = service_mod.comtypes
        orig_focus = service_mod.FocusChangedEventHandler
        orig_struct = service_mod.StructureChangedEventHandler
        orig_prop = service_mod.PropertyChangedEventHandler

        service_mod.comtypes = mock_comtypes
        service_mod.FocusChangedEventHandler = MagicMock(return_value=MagicMock())
        service_mod.StructureChangedEventHandler = MagicMock(return_value=MagicMock())
        service_mod.PropertyChangedEventHandler = MagicMock(return_value=MagicMock())

        wd.is_running.set()
        try:
            wd._run()
        finally:
            service_mod.comtypes = orig_comtypes
            service_mod.FocusChangedEventHandler = orig_focus
            service_mod.StructureChangedEventHandler = orig_struct
            service_mod.PropertyChangedEventHandler = orig_prop

        return mock_comtypes

    def test_run_registers_focus_handler_when_callback_set(self):
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._focus_callback = MagicMock()

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.AddFocusChangedEventHandler.assert_called_once()

    def test_run_deregisters_focus_handler_when_callback_cleared(self):
        """When _focus_handler is set but _focus_callback is None, handler is removed."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()

        existing_handler = MagicMock()
        wd._focus_handler = existing_handler
        wd._focus_callback = None  # cleared

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemoveFocusChangedEventHandler.assert_called_once_with(
            existing_handler
        )
        assert wd._focus_handler is None

    def test_run_registers_structure_handler_when_callback_set(self):
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._structure_callback = MagicMock()

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.AddStructureChangedEventHandler.assert_called_once()

    def test_run_registers_property_handler_with_default_ids(self):
        """When no property_ids configured, the loop uses the four default IDs."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._property_callback = MagicMock()
        wd._property_ids = None  # use defaults

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        add_call = fake_uia_client.IUIAutomation.AddPropertyChangedEventHandler.call_args
        assert add_call is not None
        prop_ids_arg = add_call[0][-1]  # last positional arg is property IDs list
        assert 30005 in prop_ids_arg
        assert 30045 in prop_ids_arg
        assert 30093 in prop_ids_arg
        assert 30128 in prop_ids_arg

    def test_run_deregisters_structure_handler_on_config_change(self):
        """Changing the structure element triggers deregister of the old handler."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()

        old_element = MagicMock(name="old_elem")
        new_element = MagicMock(name="new_elem")
        existing_handler = MagicMock()

        wd._structure_handler = existing_handler
        wd._active_structure_element = old_element
        wd._structure_callback = MagicMock()
        wd._structure_element = new_element  # config changed

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        # The loop deregisters the old handler due to config change (and may
        # re-register on new_element then deregister again in the finally block).
        # Verify the deregistration of the old element happened at least once.
        fake_uia_client.IUIAutomation.RemoveStructureChangedEventHandler.assert_any_call(
            old_element, existing_handler
        )

    def test_run_calls_comtypes_coinitialize_and_couninitialize(self):
        """The event loop must call CoInitialize and CoUninitialize on its thread."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()

        mock_comtypes = self._run_one_iteration(wd, service_mod, comtypes_stub)

        mock_comtypes.CoInitialize.assert_called_once()
        mock_comtypes.CoUninitialize.assert_called_once()

    def test_run_cleans_up_focus_handler_on_exception(self):
        """If PumpEvents raises, the finally block removes the focus handler."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()

        focus_handler = MagicMock()
        wd._focus_handler = focus_handler

        mock_comtypes = MagicMock()
        mock_comtypes.CoInitialize = MagicMock()
        mock_comtypes.CoUninitialize = MagicMock()
        mock_comtypes.client = comtypes_stub.client

        # PumpEvents raises, triggering the outer except and then finally
        comtypes_stub.client.PumpEvents.side_effect = RuntimeError("crash")

        orig_comtypes = service_mod.comtypes
        orig_focus = service_mod.FocusChangedEventHandler
        orig_struct = service_mod.StructureChangedEventHandler
        orig_prop = service_mod.PropertyChangedEventHandler

        service_mod.comtypes = mock_comtypes
        service_mod.FocusChangedEventHandler = MagicMock(return_value=MagicMock())
        service_mod.StructureChangedEventHandler = MagicMock(return_value=MagicMock())
        service_mod.PropertyChangedEventHandler = MagicMock(return_value=MagicMock())

        wd.is_running.set()
        try:
            wd._run()  # Must not propagate the RuntimeError
        finally:
            service_mod.comtypes = orig_comtypes
            service_mod.FocusChangedEventHandler = orig_focus
            service_mod.StructureChangedEventHandler = orig_struct
            service_mod.PropertyChangedEventHandler = orig_prop

        fake_uia_client.IUIAutomation.RemoveFocusChangedEventHandler.assert_called_once_with(
            focus_handler
        )
        mock_comtypes.CoUninitialize.assert_called_once()

    # --- Exception paths: handler add/remove failures (lines 97-195) ---

    def test_add_focus_handler_exception_is_caught(self):
        """AddFocusChangedEventHandler raises → exception logged, handler stays None."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._focus_callback = MagicMock()
        fake_uia_client.IUIAutomation.AddFocusChangedEventHandler.side_effect = OSError("COM fail")

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        # Handler should remain None because add failed
        assert wd._focus_handler is None

    def test_remove_focus_handler_exception_is_caught(self):
        """RemoveFocusChangedEventHandler raises → exception logged, handler set to None."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        existing_handler = MagicMock()
        wd._focus_handler = existing_handler
        wd._focus_callback = None  # want to remove
        fake_uia_client.IUIAutomation.RemoveFocusChangedEventHandler.side_effect = OSError("fail")

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        # Handler should be cleared even though remove raised
        assert wd._focus_handler is None

    def test_add_structure_handler_exception_is_caught(self):
        """AddStructureChangedEventHandler raises → exception logged, no handler set."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._structure_callback = MagicMock()
        fake_uia_client.IUIAutomation.AddStructureChangedEventHandler.side_effect = OSError(
            "COM fail"
        )

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        # Handler stays None because add failed
        assert wd._structure_handler is None

    def test_remove_structure_handler_exception_is_caught(self):
        """RemoveStructureChangedEventHandler raises during disable → caught, handler cleared."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        existing_handler = MagicMock()
        wd._structure_handler = existing_handler
        wd._active_structure_element = MagicMock()
        wd._structure_callback = None  # want to remove
        fake_uia_client.IUIAutomation.RemoveStructureChangedEventHandler.side_effect = OSError(
            "fail"
        )

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        assert wd._structure_handler is None
        assert wd._active_structure_element is None

    def test_add_property_handler_exception_is_caught(self):
        """AddPropertyChangedEventHandler raises → exception logged, no handler set."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._property_callback = MagicMock()
        fake_uia_client.IUIAutomation.AddPropertyChangedEventHandler.side_effect = OSError(
            "COM fail"
        )

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        assert wd._property_handler is None

    def test_remove_property_handler_exception_is_caught(self):
        """RemovePropertyChangedEventHandler raises during disable → caught, handler cleared."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        existing_handler = MagicMock()
        wd._property_handler = existing_handler
        wd._active_property_element = MagicMock()
        wd._active_property_ids = [30005]
        wd._property_callback = None  # want to remove
        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.side_effect = OSError(
            "fail"
        )

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        assert wd._property_handler is None
        assert wd._active_property_element is None
        assert wd._active_property_ids is None

    # --- Exception paths: finally block cleanup (lines 203-238) ---

    def _run_with_crash(self, wd, service_mod, comtypes_stub):
        """Run _run() with PumpEvents raising, then return mock_comtypes for assertions."""
        mock_comtypes = MagicMock()
        mock_comtypes.CoInitialize = MagicMock()
        mock_comtypes.CoUninitialize = MagicMock()
        mock_comtypes.client = comtypes_stub.client

        comtypes_stub.client.PumpEvents.side_effect = RuntimeError("crash")

        orig_comtypes = service_mod.comtypes
        orig_focus = service_mod.FocusChangedEventHandler
        orig_struct = service_mod.StructureChangedEventHandler
        orig_prop = service_mod.PropertyChangedEventHandler

        service_mod.comtypes = mock_comtypes
        service_mod.FocusChangedEventHandler = MagicMock(return_value=MagicMock())
        service_mod.StructureChangedEventHandler = MagicMock(return_value=MagicMock())
        service_mod.PropertyChangedEventHandler = MagicMock(return_value=MagicMock())

        wd.is_running.set()
        try:
            wd._run()
        finally:
            service_mod.comtypes = orig_comtypes
            service_mod.FocusChangedEventHandler = orig_focus
            service_mod.StructureChangedEventHandler = orig_struct
            service_mod.PropertyChangedEventHandler = orig_prop

        return mock_comtypes

    def test_finally_cleans_up_structure_handler_on_crash(self):
        """On PumpEvents crash, finally block removes structure handler."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        structure_handler = MagicMock()
        active_element = MagicMock()
        wd._structure_handler = structure_handler
        wd._active_structure_element = active_element

        self._run_with_crash(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemoveStructureChangedEventHandler.assert_called_with(
            active_element, structure_handler
        )
        assert wd._structure_handler is None
        assert wd._active_structure_element is None

    def test_finally_cleans_up_property_handler_on_crash(self):
        """On PumpEvents crash, finally block removes property handler."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        property_handler = MagicMock()
        active_element = MagicMock()
        wd._property_handler = property_handler
        wd._active_property_element = active_element
        wd._active_property_ids = [30005]

        self._run_with_crash(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.assert_called_with(
            active_element, property_handler
        )
        assert wd._property_handler is None
        assert wd._active_property_element is None
        assert wd._active_property_ids is None

    def test_finally_focus_remove_exception_swallowed(self):
        """If RemoveFocusChangedEventHandler raises in finally, exception is swallowed.

        We set _focus_callback so the loop body doesn't enter the 'elif remove' branch.
        The pre-existing _focus_handler + callback means the 'if add' branch is also
        skipped (handler already exists). PumpEvents crashes → finally tries Remove → raises.
        """
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._focus_callback = MagicMock()  # keeps elif-remove from triggering in loop
        wd._focus_handler = MagicMock()  # already exists, so if-add is skipped
        fake_uia_client.IUIAutomation.RemoveFocusChangedEventHandler.side_effect = OSError(
            "COM dead"
        )

        self._run_with_crash(wd, service_mod, comtypes_stub)

        # Should not propagate, handler cleared despite exception
        assert wd._focus_handler is None

    def test_finally_structure_remove_exception_swallowed(self):
        """If RemoveStructureChangedEventHandler raises in finally, exception is swallowed."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        element = MagicMock()
        wd._structure_callback = MagicMock()
        wd._structure_element = element
        wd._structure_handler = MagicMock()
        wd._active_structure_element = element
        fake_uia_client.IUIAutomation.RemoveStructureChangedEventHandler.side_effect = OSError(
            "COM dead"
        )

        self._run_with_crash(wd, service_mod, comtypes_stub)

        assert wd._structure_handler is None
        assert wd._active_structure_element is None

    def test_finally_property_remove_exception_swallowed(self):
        """If RemovePropertyChangedEventHandler raises in finally, exception is swallowed."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        element = MagicMock()
        wd._property_callback = MagicMock()
        wd._property_element = element
        wd._property_ids = [30005]
        wd._property_handler = MagicMock()
        wd._active_property_element = element
        wd._active_property_ids = [30005]
        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.side_effect = OSError(
            "COM dead"
        )

        self._run_with_crash(wd, service_mod, comtypes_stub)

        assert wd._property_handler is None
        assert wd._active_property_element is None
        assert wd._active_property_ids is None

    # --- GetRootElement fallback paths (lines 121, 135, 159, 175) ---

    def test_structure_remove_uses_root_element_when_no_active_element(self):
        """When active_structure_element is None, removal uses GetRootElement."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        existing_handler = MagicMock()
        root = MagicMock(name="root")
        fake_uia_client.IUIAutomation.GetRootElement.return_value = root

        wd._structure_handler = existing_handler
        wd._active_structure_element = None  # triggers GetRootElement fallback
        wd._structure_callback = None  # want to remove

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemoveStructureChangedEventHandler.assert_any_call(
            root, existing_handler
        )

    def test_property_remove_uses_root_element_when_no_active_element(self):
        """When active_property_element is None, removal uses GetRootElement."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        existing_handler = MagicMock()
        root = MagicMock(name="root")
        fake_uia_client.IUIAutomation.GetRootElement.return_value = root

        wd._property_handler = existing_handler
        wd._active_property_element = None  # triggers GetRootElement fallback
        wd._active_property_ids = [30005]
        wd._property_callback = None  # want to remove

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.assert_any_call(
            root, existing_handler
        )

    def test_structure_add_uses_root_element_when_no_element_specified(self):
        """When _structure_element is None, registration uses GetRootElement."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        root = MagicMock(name="root")
        fake_uia_client.IUIAutomation.GetRootElement.return_value = root

        wd._structure_callback = MagicMock()
        wd._structure_element = None  # triggers GetRootElement

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        call_args = fake_uia_client.IUIAutomation.AddStructureChangedEventHandler.call_args
        assert call_args is not None
        assert call_args[0][0] is root

    def test_property_add_uses_root_element_when_no_element_specified(self):
        """When _property_element is None, registration uses GetRootElement."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        root = MagicMock(name="root")
        fake_uia_client.IUIAutomation.GetRootElement.return_value = root

        wd._property_callback = MagicMock()
        wd._property_element = None  # triggers GetRootElement

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        call_args = fake_uia_client.IUIAutomation.AddPropertyChangedEventHandler.call_args
        assert call_args is not None
        assert call_args[0][0] is root

    # --- Property handler with custom IDs (line 181-184) ---

    def test_property_handler_with_custom_ids(self):
        """When custom property_ids are set, they're passed instead of defaults."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        wd._property_callback = MagicMock()
        wd._property_ids = [30005, 30045]  # custom IDs

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        add_call = fake_uia_client.IUIAutomation.AddPropertyChangedEventHandler.call_args
        assert add_call is not None
        prop_ids_arg = add_call[0][-1]
        assert prop_ids_arg == [30005, 30045]

    # --- Property deregistration on config change (lines 148-168) ---

    def test_property_deregisters_on_element_change(self):
        """Changing property_element triggers deregister of old + register of new."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        old_element = MagicMock(name="old")
        new_element = MagicMock(name="new")
        existing_handler = MagicMock()

        wd._property_handler = existing_handler
        wd._active_property_element = old_element
        wd._active_property_ids = [30005]
        wd._property_callback = MagicMock()
        wd._property_element = new_element  # config changed

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.assert_any_call(
            old_element, existing_handler
        )

    def test_property_deregisters_on_ids_change(self):
        """Changing property_ids triggers deregister of old + register of new."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        element = MagicMock()
        existing_handler = MagicMock()

        wd._property_handler = existing_handler
        wd._active_property_element = element
        wd._active_property_ids = [30005]
        wd._property_callback = MagicMock()
        wd._property_element = element  # same element
        wd._property_ids = [30005, 30045]  # different IDs → config changed

        self._run_one_iteration(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.assert_any_call(
            element, existing_handler
        )

    # --- Finally block with no active element (GetRootElement fallback) ---

    def test_finally_structure_uses_root_when_no_active_element(self):
        """In finally cleanup, if active_structure_element is None, uses GetRootElement."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        root = MagicMock(name="root")
        fake_uia_client.IUIAutomation.GetRootElement.return_value = root

        structure_handler = MagicMock()
        wd._structure_handler = structure_handler
        wd._active_structure_element = None

        self._run_with_crash(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemoveStructureChangedEventHandler.assert_called_with(
            root, structure_handler
        )

    def test_finally_property_uses_root_when_no_active_element(self):
        """In finally cleanup, if active_property_element is None, uses GetRootElement."""
        wd, fake_uia_client, service_mod, comtypes_stub = self._setup()
        root = MagicMock(name="root")
        fake_uia_client.IUIAutomation.GetRootElement.return_value = root

        property_handler = MagicMock()
        wd._property_handler = property_handler
        wd._active_property_element = None
        wd._active_property_ids = [30005]

        self._run_with_crash(wd, service_mod, comtypes_stub)

        fake_uia_client.IUIAutomation.RemovePropertyChangedEventHandler.assert_called_with(
            root, property_handler
        )
