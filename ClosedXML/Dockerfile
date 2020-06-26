FROM mcr.microsoft.com/dotnet/core/sdk:3.0 as build

ARG projname=ClosedXML
ARG key=TODO
ARG packageversion=1.0.1
ARG nugetserver=http://nuget.dev.logicsoftware.net

WORKDIR /$projname

COPY $projname/$projname.csproj $projname/
COPY $projname/$projname.nuspec $projname/

RUN dotnet restore $projname/$projname.csproj

COPY $projname $projname/

WORKDIR /$projname/$projname
RUN dotnet pack /p:Version=$packageversion -o packages
RUN dotnet nuget push packages/LogicSoftware.$projname.$packageversion.nupkg -k $key -s $nugetserver