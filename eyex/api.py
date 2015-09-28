from collections import namedtuple
import ctypes as c
import eyex.types as tx

SampleGaze = namedtuple('SampleGaze', ['data_mode', 'timestamp', 'x', 'y'])
SampleFixation = namedtuple('SampleFixation', ['data_mode', 'event_type', 'timestamp', 'x', 'y'])


class EyeXInterface():
    on_event = []

    def __init__(self, lib_location='Tobii.EyeX.Client.dll', fixation=False):

        self.eyex_dll = c.cdll.LoadLibrary(lib_location)
        self.fixation = fixation

        self.latest_sample = None

        self.interactor_snapshot = c.c_voidp()
        self.context = c.c_voidp()
        self.interactor_id = "3425g"

        event_handler_ticket = c.c_int(0)
        connection_state_changed_ticket = c.c_int(0)

        self._c_event_handler = tx.EVENT_HANDLER(self._event_handler)
        self._c_on_snapshot_committed = tx.ON_SNAPSHOT_COMMITTED(self._on_snapshot_committed)
        self._c_connection_handler = tx.CONNECTION_HANDLER(self._connection_handler)

        ret = self.eyex_dll.txInitializeEyeX(tx.TX_SYSTEMCOMPONENTOVERRIDEFLAGS.TX_SYSTEMCOMPONENTOVERRIDEFLAG_NONE,
                                               None, None, None)

        ret = self.eyex_dll.txCreateContext(c.byref(self.context), tx.TX_FALSE)
        self._initialize_interactor_snapshot()
        ret = self.eyex_dll.txRegisterConnectionStateChangedHandler(self.context,
                                                                    c.byref(connection_state_changed_ticket),
                                                                    self._c_connection_handler, None)

        ret = self.eyex_dll.txRegisterEventHandler(self.context, c.byref(event_handler_ticket), self._c_event_handler,
                                                   None)
        ret = self.eyex_dll.txEnableConnection(self.context)

    def __del__(self):
        if hasattr(self, 'eyex_dll') and self.eyex_dll is not None:
            ret = self.eyex_dll.txDisableConnection(self.context)
            ret = self.eyex_dll.txShutdownContext(self.context, 500, tx.TX_FALSE)
            ret = self.eyex_dll.txReleaseContext(c.byref(self.context))

    def _initialize_interactor_snapshot(self):
        interactor = c.c_voidp()

        if self.fixation:
            params = tx.TX_FIXATIONDATAPARAMS(tx.TX_FIXATIONDATAMODE_SLOW)
        else:
            params = tx.TX_GAZEPOINTDATAPARAMS(tx.TX_GAZEPOINTDATAMODE_LIGHTLYFILTERED)

        ret = self.eyex_dll.txCreateGlobalInteractorSnapshot(self.context, self.interactor_id,
                                                             c.byref(self.interactor_snapshot),
                                                             c.byref(interactor))

        if self.fixation:
            ret = self.eyex_dll.txCreateFixationDataBehavior(interactor,  c.byref(params))
        else:
            ret = self.eyex_dll.txCreateGazePointDataBehavior(interactor, c.byref(params))

        ret = self.eyex_dll.txReleaseObject(c.byref(interactor))


    def _event_handler(self, async_data, userParam):
        event = c.c_voidp()
        behavior = c.c_voidp()

        self.eyex_dll.txGetAsyncDataContent(async_data, c.byref(event))

        if self.fixation:
            if self.eyex_dll.txGetEventBehavior(event, c.byref(behavior),
                tx.TX_BEHAVIORTYPE_FIXATIONDATA) == tx.TX_RESULT_OK:
                event_params = tx.TX_FIXATIONDATAEVENTPARAMS()
                if self.eyex_dll.txGetFixationDataEventParams(behavior, c.byref(event_params)) == tx.TX_RESULT_OK:
                    sample = SampleFixation(int(event_params.FixationDataMode),
                                    int(event_params.FixationDataEventType),
                                    float(event_params.timestamp),
                                    float(event_params.x),
                                    float(event_params.y))
                    self.latest_sample = sample
                    for callback in self.on_event:
                        callback(sample)
                self.eyex_dll.txReleaseObject(c.byref(behavior))
        else:
            if self.eyex_dll.txGetEventBehavior(event, c.byref(behavior),
                tx.TX_BEHAVIORTYPE_GAZEPOINTDATA) == tx.TX_RESULT_OK:
                event_params = tx.TX_GAZEPOINTDATAEVENTPARAMS()
                if self.eyex_dll.txGetGazePointDataEventParams(behavior, c.byref(event_params)) == tx.TX_RESULT_OK:
                    sample = SampleGaze(int(event_params.GazePointDataMode),
                                    float(event_params.timestamp),
                                    float(event_params.x),
                                    float(event_params.y))
                    self.latest_sample = sample
                    for callback in self.on_event:
                        callback(sample)
                self.eyex_dll.txReleaseObject(c.byref(behavior))

        self.eyex_dll.txReleaseObject(c.byref(event))

    def _on_snapshot_committed(self, async_data, param):
        result = c.c_int(0)
        self.eyex_dll.txGetAsyncDataResultCode(async_data, c.pointer(result))

    def _connection_handler(self, connection_state, user_param):
        if connection_state == tx.TX_CONNECTIONSTATE.TX_CONNECTIONSTATE_CONNECTED:
            ret = self.eyex_dll.txCommitSnapshotAsync(self.interactor_snapshot, self._c_on_snapshot_committed, None)

