# csipopt

Copyright (c) 2010-2013 Anders Gustafsson, Cureos AB.<br/>
Published under the Eclipse Public License.

## Introduction

*csipopt* provides a small and simple C# interface to the *Ipopt* non-linear optimizer. If the interface is
contained in a .NET class library, the interface is accessible via any .NET language. *.NET Framework 2.0*
or higher is required.

It is also possible to build and use *csipopt* in *Silverlight* applications using elevated trust privileges 
and *Windows Store* applications.

The most up-to-date usage information can be found on the [wiki pages](https://github.com/cureos/csipopt/wiki).

## Prerequisites

To successfully run an application calling the C# *Ipopt* interface, a dynamic linked library of *Ipopt* is 
additionally required. Pre-compiled libraries can be downloaded from the 
[following location](http://www.coin-or.org/download/binary/Ipopt/).

At the time of this writing, the most up-to-date binaries, version 3.11.0, can be downloaded 
[here](http://www.coin-or.org/download/binary/Ipopt/Ipopt-3.11.0-Win32-Win64-dll.7z).

Download binaries suitable for your platform, unpack the compressed *Ipopt* binary archive, from the folder 
*./lib/PLATFORM/ReleaseMKL* (where *PLATFORM* is *win32* or *x64*) copy the DLL files and place them in a folder 
that is accessible from the application or system path.

(Note! The actual name of the DLL is of course dependent upon which version of *Ipopt* binaries you are 
downloading.)


## Usage

It is possible to build stand-alone class libraries with all *csipopt* code from the *Cureos.Numerics* solution. 
Open this solution in Visual Studio 2010 or 2012 and build the *Cureos.Numerics* (.NET 4 and higher) or *Metro.Cureos.Numerics* 
(*Windows Store* applications) project, then reference the built project in your own application.

To use the C# interface in your own code, you can also copy the C# files _Ipopt*.cs_ to your class library. 
(Potentially, you may need to redefine the `IpoptDllName` in the *IpoptAdapter.cs* file to match the name of the 
native *Ipopt* DLL.) 

A number of build scripts are provided in the build sub-folder to demonstrate example usage of the C# *Ipopt* interface.


### Visual Studio examples

When running from the Visual Studio command prompt, create a class library by running the batch script

    build_library

Create an example application from the C# file *Program.cs* by running the batch script

    build_hs071_cs

There is an alternative C# build script enforcing the intermediate callback method (introduced in 
*Ipopt 3.9*) to be utilized. To create an example application using the intermediate callback method, run

    build_hs071_cs_with_intermediate

Analogously, create an example application from the Visual Basic file *Program.vb* or F# file *Program.fs*
by running either of the batch scripts

    build_hs071_vb
    build_hs071_fs

To run the respective applications, type

    hs071_cs
    hs071_cs_intermediate
    hs071_vb
    hs071_fs

### Mono examples
 
The interface is also applicable to *Mono*, and can be compiled using the *gmcs* compiler.
When running from the *Mono* command prompt, create a class library by running the batch script

    build_library_with_mono

Create an example application from the C# file *Program.cs* by running the batch script

    build_hs071_cs_with_mono

To run the example application, type

    mono hs071_cs_mono.exe

NOTE! The *Mono* runtime is more conservative when handling array marshaling in unmanaged
function pointers, i.e. the callback functions. This is handled through a modification of
the signatures of the `Eval_Jac_G_CB` and `Eval_H_CB` delegates and explicit copying of arrays
in the managed function wrappers in `IpoptProblem`. To access these modifications, build the
library with the `MONO` symbol defined.


## Revision

Last updated on May 27, 2013 by Anders Gustafsson, anders[at]cureos[dot]com, http://www.cureos.com.
