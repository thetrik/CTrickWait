#CTrickWait

## Asynchronous waiting for kernel objects



Hello everyone.



This class allows to wait for Windows kernel objects and generate an event when an object switches to the signaled state or a timeout has elapsed.

The class has 3 methods: **vbWaitForSingleObject**, **vbWaitForMultipleObjects** and **Abort**. The first two methods are the analogs of the corresponding WINAPI functions [WaitForSingleObject](https://docs.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-waitforsingleobject) and [WaitForMultipleObjects](https://docs.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-waitformultipleobjects).



 As soon as an object (or all the objects) changes the state to signaled the event **OnWait** is fired. The arguments of the events contains the event handle (or the pointer to the handles) and the returned value. **Abort** method allows to break any pending waiting operation. It can either returns immediately or wait until the request will be processed.

The class also contains property **IsActive** which shows if there is an active waiting operation.

There are 3 examples of class usage: waiting for a waitable timer, waiting for a process completion and monitiring file operations.

## How does this work?

The class has an assembly thunk which creates a thread and runs waiting in this thread. All the requests are managed by an event handle and the results are transmitted to the main thread using windows messages. It also uses APC requests to abort the waiting operations. I hope the class are quite safe for IDE debugging so you can use the Stop button or End operator. To achive this it uses a simple COM object which controls the lifetime of class. When you use the Stop button or End statement the runtime doesn't call the Terminate event of classes but it always releases all the resources. A class instance always hold the special object and if this object is released the ASM thunk uninitializes the thread. 

The module is poorly tested so bugs are possible. I would be very glad to any bug-reports, wherever possible I will correct them.


Thank you all for attention!



Best Regards,



The trick.