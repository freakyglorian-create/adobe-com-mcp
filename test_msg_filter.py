"""Implement IMessageFilter for automatic COM retry on RPC_E_RETRYLATER."""
import pythoncom
import win32com.client
import ctypes
from ctypes import wintypes

# Constants
PENDINGMSG_WAITDEFPROCESS = 2
PENDINGMSG_WAITNOPROCESS = 1
PENDINGMSG_WAITALL = 3
PENDINGMSG_CANCELCALL = 0

SERVERCALL_ISHANDLED = 0
SERVERCALL_REJECTED = 1
SERVERCALL_RETRYLATER = 2

class MessageFilter:
    """COM message filter that automatically retries calls rejected with RPC_E_RETRYLATER."""
    _com_interfaces_ = []
    
    def __init__(self, max_retries=50, retry_delay_ms=200):
        self.max_retries = max_retries
        self.retry_delay_ms = retry_delay_ms
        self._retry_count = 0
    
    def HandleInComingCall(self, dwCallType, htaskCaller, dwTickCount, lpInterfaceInfo):
        return SERVERCALL_ISHANDLED
    
    def RetryRejectedCall(self, htaskCallee, dwTickCount, dwRejectType):
        if dwRejectType == SERVERCALL_RETRYLATER:
            self._retry_count += 1
            if self._retry_count <= self.max_retries:
                # Sleep and retry
                ctypes.windll.kernel32.Sleep(self.retry_delay_ms)
                return 1  # retry immediately (value > 0 means retry after N ms, but we already slept)
            self._retry_count = 0
            return -1  # cancel
        return -1  # cancel for other rejection types
    
    def MessagePending(self, htaskCallee, dwTickCount, dwPendingType):
        return PENDINGMSG_WAITDEFPROCESS


def register_message_filter(max_retries=50, retry_delay_ms=200):
    """Register a COM message filter for the current thread. Returns the previous filter."""
    mf = MessageFilter(max_retries, retry_delay_ms)
    # Use pythoncom's message filter registration
    try:
        prev = pythoncom.CoRegisterMessageFilter(mf)
        return prev
    except Exception:
        # Fallback: implement via ctypes
        pass
    return None


# Test
if __name__ == "__main__":
    import time
    
    pythoncom.CoInitialize()
    print("Registering message filter...")
    register_message_filter(max_retries=30, retry_delay_ms=200)
    
    print("Connecting to PS...")
    app = win32com.client.GetActiveObject("Photoshop.Application")
    print(f"Version: {app.Version}")
    
    # Run several operations to test
    print("\nRunning test operations...")
    for i in range(5):
        try:
            r = app.DoJavaScript(f'"test-{i}"')
            print(f"  Op {i+1}: {r}")
        except Exception as e:
            print(f"  Op {i+1} FAILED: {e}")
    
    print("\nDone")
