FROM mcr.microsoft.com/dotnet/core/sdk:3.0 as build

ARG projname=ClosedXML
ARG key=TODO
ARG packageversion=1.0.1
ARG nugetserver=http://nuget.dev.logicsoftware.net

WORKDIR /$projname

COPY $projname/$projname.csproj $projname/
COPY $projname/$projname.nuspec $projname/

RUN dotnet restore $projname/$projname.csproj -c Release.Signed

COPY $projname $projname/

WORKDIR /$projname/$projname
RUN dotnet pack /p:Version=$packageversion -o packages -c Release.Signed
RUN dotnet nuget push packages/$projname.$packageversion.nupkg -k $key -s $nugetserver