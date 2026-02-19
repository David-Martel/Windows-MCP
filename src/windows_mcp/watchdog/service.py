"""
Unified WatchDog Service for monitoring UI Automation events.
Allows single instantiation to handle multiple monitors (Focus, Structure) safely in one STA thread.
"""

import logging
from threading import Event, Thread

import comtypes
import comtypes.client

from windows_mcp.uia.core import _AutomationClient
from windows_mcp.uia.enums import PropertyId, TreeScope

from .event_handlers import (
    AutomationEventHandler,
    FocusChangedEventHandler,
    PropertyChangedEventHandler,
    StructureChangedEventHandler,
)

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


class WatchDog:
    def __init__(self):
        self.uia_client = _AutomationClient.instance()
        self.uia = self.uia_client.IUIAutomation
        self.is_running = Event()
        self.thread = None

        # Callbacks
        self._focus_callback = None
        self._structure_callback = None
        self._structure_element = None
        self._property_callback = None
        self._property_element = None
        self._property_ids = None
        self._automation_callback = None
        self._automation_event_id = None
        self._automation_element = None

        # Internal state for tracking active handlers
        self._focus_handler = None
        self._structure_handler = None
        self._active_structure_element = None
        self._property_handler = None
        self._active_property_element = None
        self._active_property_ids = None
        self._automation_handler = None
        self._active_automation_event_id = None
        self._active_automation_element = None

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()

    def start(self):
        """Start the watchdog service thread."""
        if self.is_running.is_set():
            return
        self.is_running.set()
        self.thread = Thread(target=self._run, name="WatchDogThread")
        self.thread.daemon = True
        self.thread.start()

    def stop(self):
        """Stop the watchdog service thread."""
        if not self.is_running.is_set():
            return
        self.is_running.clear()
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=2.0)

    def set_focus_callback(self, callback):
        """Set the callback for focus changes. Pass None to disable."""
        self._focus_callback = callback

    def set_structure_callback(self, callback, element=None):
        """Set the callback for structure changes. Pass None to disable.
        Optionally specify an element to watch (defaults to RootElement)."""
        self._structure_callback = callback
        self._structure_element = element

    def set_property_callback(self, callback, element=None, property_ids=None):
        """Set the callback for property changes. Pass None to disable.
        Optionally specify an element to watch (defaults to RootElement)
        and a list of property IDs to monitor."""
        self._property_callback = callback
        self._property_element = element
        self._property_ids = property_ids

    def set_automation_callback(self, callback, event_id=None, element=None):
        """Set the callback for generic automation events. Pass None to disable.
        ``event_id`` is a UIA event constant (e.g. ``EventId.UIA_Window_WindowOpenedEventId``).
        ``element`` scopes the subscription (defaults to RootElement)."""
        self._automation_callback = callback
        self._automation_event_id = event_id
        self._automation_element = element

    def _run(self):
        """Main event loop running in a dedicated STA thread."""
        comtypes.CoInitialize()
        try:
            while self.is_running.is_set():
                # --- Focus Monitoring ---
                if self._focus_callback and not self._focus_handler:
                    try:
                        self._focus_handler = FocusChangedEventHandler(self)
                        self.uia.AddFocusChangedEventHandler(None, self._focus_handler)
                    except Exception as e:
                        logger.debug("Failed to add focus handler: %s", e)
                elif not self._focus_callback and self._focus_handler:
                    try:
                        self.uia.RemoveFocusChangedEventHandler(self._focus_handler)
                    except Exception as e:
                        logger.debug("Failed to remove focus handler: %s", e)
                    self._focus_handler = None

                # --- Structure Monitoring ---
                # Check if we need to UNREGISTER because configuration changed or disabled
                config_changed = self._structure_element != self._active_structure_element

                should_be_active = self._structure_callback is not None
                is_active = self._structure_handler is not None

                if is_active and (not should_be_active or config_changed):
                    try:
                        target = (
                            self._active_structure_element
                            if self._active_structure_element
                            else self.uia.GetRootElement()
                        )
                        self.uia.RemoveStructureChangedEventHandler(target, self._structure_handler)
                    except Exception as e:
                        logger.debug("Failed to remove structure handler: %s", e)
                    self._structure_handler = None
                    self._active_structure_element = None
                    is_active = False

                if should_be_active and not is_active:
                    try:
                        target = (
                            self._structure_element
                            if self._structure_element
                            else self.uia.GetRootElement()
                        )
                        scope = TreeScope.TreeScope_Subtree

                        self._structure_handler = StructureChangedEventHandler(self)
                        self.uia.AddStructureChangedEventHandler(
                            target, scope, None, self._structure_handler
                        )
                        self._active_structure_element = target
                    except Exception as e:
                        logger.debug("Failed to add structure handler: %s", e)

                # --- Property Monitoring ---
                config_changed = (self._property_element != self._active_property_element) or (
                    self._property_ids != self._active_property_ids
                )

                should_be_active = self._property_callback is not None
                is_active = self._property_handler is not None

                if is_active and (not should_be_active or config_changed):
                    try:
                        target = (
                            self._active_property_element
                            if self._active_property_element
                            else self.uia.GetRootElement()
                        )
                        self.uia.RemovePropertyChangedEventHandler(target, self._property_handler)
                    except Exception as e:
                        logger.debug("Failed to remove property handler: %s", e)
                    self._property_handler = None
                    self._active_property_element = None
                    self._active_property_ids = None
                    is_active = False

                if should_be_active and not is_active:
                    try:
                        target = (
                            self._property_element
                            if self._property_element
                            else self.uia.GetRootElement()
                        )
                        scope = TreeScope.TreeScope_Subtree

                        # Monitor common properties if none specified
                        p_ids = (
                            self._property_ids
                            if self._property_ids
                            else [
                                PropertyId.NameProperty,
                                PropertyId.ValueValueProperty,
                                PropertyId.LegacyIAccessibleValueProperty,
                                PropertyId.ToggleToggleStateProperty,
                            ]
                        )

                        self._property_handler = PropertyChangedEventHandler(self)
                        self.uia.AddPropertyChangedEventHandler(
                            target, scope, None, self._property_handler, p_ids
                        )

                        self._active_property_element = target
                        self._active_property_ids = p_ids
                    except Exception as e:
                        logger.debug("Failed to add property handler: %s", e)

                # --- Automation Event Monitoring ---
                auto_config_changed = (
                    self._automation_event_id != self._active_automation_event_id
                ) or (self._automation_element != self._active_automation_element)

                should_be_active = self._automation_callback is not None
                is_active = self._automation_handler is not None

                if is_active and (not should_be_active or auto_config_changed):
                    try:
                        target = (
                            self._active_automation_element
                            if self._active_automation_element
                            else self.uia.GetRootElement()
                        )
                        self.uia.RemoveAutomationEventHandler(
                            self._active_automation_event_id, target, self._automation_handler
                        )
                    except Exception as e:
                        logger.debug("Failed to remove automation handler: %s", e)
                    self._automation_handler = None
                    self._active_automation_event_id = None
                    self._active_automation_element = None
                    is_active = False

                if should_be_active and not is_active and self._automation_event_id is not None:
                    try:
                        target = (
                            self._automation_element
                            if self._automation_element
                            else self.uia.GetRootElement()
                        )
                        scope = TreeScope.TreeScope_Subtree
                        self._automation_handler = AutomationEventHandler(self)
                        self.uia.AddAutomationEventHandler(
                            self._automation_event_id, target, scope, None,
                            self._automation_handler,
                        )
                        self._active_automation_event_id = self._automation_event_id
                        self._active_automation_element = target
                    except Exception as e:
                        logger.debug("Failed to add automation handler: %s", e)

                # Pump events for this thread
                comtypes.client.PumpEvents(0.1)

        except Exception as e:
            logger.error("WatchDogService died: %s", e)
        finally:
            # Cleanup handlers on exit
            if self._focus_handler:
                try:
                    self.uia.RemoveFocusChangedEventHandler(self._focus_handler)
                except Exception:
                    pass
                self._focus_handler = None

            if self._structure_handler:
                try:
                    target = (
                        self._active_structure_element
                        if self._active_structure_element
                        else self.uia.GetRootElement()
                    )
                    self.uia.RemoveStructureChangedEventHandler(target, self._structure_handler)
                except Exception:
                    pass
                self._structure_handler = None
                self._active_structure_element = None

            if self._property_handler:
                try:
                    target = (
                        self._active_property_element
                        if self._active_property_element
                        else self.uia.GetRootElement()
                    )
                    self.uia.RemovePropertyChangedEventHandler(target, self._property_handler)
                except Exception:
                    pass
                self._property_handler = None
                self._active_property_element = None
                self._active_property_ids = None

            if self._automation_handler:
                try:
                    target = (
                        self._active_automation_element
                        if self._active_automation_element
                        else self.uia.GetRootElement()
                    )
                    self.uia.RemoveAutomationEventHandler(
                        self._active_automation_event_id, target, self._automation_handler
                    )
                except Exception:
                    pass
                self._automation_handler = None
                self._active_automation_event_id = None
                self._active_automation_element = None

            comtypes.CoUninitialize()
