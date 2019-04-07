Module OSInfo
    Private Const PRODUCT_UNDEFINED As Integer = &H0
    Private Const PRODUCT_ULTIMATE As Integer = &H1
    Private Const PRODUCT_HOME_BASIC As Integer = &H2
    Private Const PRODUCT_HOME_PREMIUM As Integer = &H3
    Private Const PRODUCT_ENTERPRISE As Integer = &H4
    Private Const PRODUCT_BUSINESS As Integer = &H6
    Private Const PRODUCT_STANDARD_SERVER As Integer = &H7
    Private Const PRODUCT_DATACENTER_SERVER As Integer = &H8
    Private Const PRODUCT_ENTERPRISE_SERVER As Integer = &HA
    Private Const PRODUCT_STARTER As Integer = &HB
    Private Const PRODUCT_WEB_SERVER As Integer = &H11
    Private Const PRODUCT_HOME_SERVER As Integer = &H13
    Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT As Integer = &H1E
    Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY As Integer = &H1F
    Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING As Integer = &H20
    Private Const PRODUCT_SERVER_FOUNDATION As Integer = &H21
    Private Const PRODUCT_HOME_PREMIUM_SERVER As Integer = &H22
    Private Const PRODUCT_PROFESSIONAL As Integer = &H30


    Private Const VER_NT_WORKSTATION As Integer = 1
    Private Const VER_NT_SERVER As Integer = 3
    Private Const VER_SUITE_ENTERPRISE As Integer = 2
    Private Const VER_SUITE_PERSONAL As Integer = 512


    <StructLayout(LayoutKind.Sequential)> _
    Private Structure OSVERSIONINFOEX
        Public dwOSVersionInfoSize As Integer
        Public dwMajorVersion As Integer
        Public dwMinorVersion As Integer
        Public dwBuildNumber As Integer
        Public dwPlatformId As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> _
        Public szCSDVersion As String
        Public wServicePackMajor As Short
        Public wServicePackMinor As Short
        Public wSuiteMask As Short
        Public wProductType As Byte
        Public wReserved As Byte
    End Structure


    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SYSTEM_INFO
        Friend uProcessorInfo As _PROCESSOR_INFO_UNION
        Public dwPageSize As UInteger
        Public lpMinimumApplicationAddress As IntPtr
        Public lpMaximumApplicationAddress As IntPtr
        Public dwActiveProcessorMask As IntPtr
        Public dwNumberOfProcessors As UInteger
        Public dwProcessorType As UInteger
        Public dwAllocationGranularity As UInteger
        Public dwProcessorLevel As UShort
        Public dwProcessorRevision As UShort
    End Structure


    <StructLayout(LayoutKind.Explicit)> _
    Private Structure _PROCESSOR_INFO_UNION
        <FieldOffset(0)> _
        Friend dwOemId As UInteger
        <FieldOffset(0)> _
        Friend wProcessorArchitecture As UShort
        <FieldOffset(2)> _
        Friend wReserved As UShort
    End Structure


    <DllImport("Kernel32.dll")> _
    Private Function GetProductInfo(ByVal osMajorVersion As Integer, ByVal osMinorVersion As Integer, ByVal spMajorVersion As Integer, ByVal spMinorVersion As Integer, ByRef edition As Integer) As Boolean
    End Function


    <DllImport("kernel32.dll")> _
    Private Function GetVersionEx(ByRef osVersionInfo As OSVERSIONINFOEX) As Boolean
    End Function


    Private m_Edition As String, m_Name As String


    Sub New()
    End Sub


    Public ReadOnly Property Bits()
        Get
            Return IntPtr.Size * 8
        End Get
    End Property


    Public ReadOnly Property Edition()
        Get
            If m_Edition <> "" Then
                Return m_Edition
            Else
                Dim tEdition As String = String.Empty
                Dim osVersion As OperatingSystem = Environment.OSVersion
                Dim osVersionInfo As New OSVERSIONINFOEX()
                osVersionInfo.dwOSVersionInfoSize = Marshal.SizeOf(GetType(OSVERSIONINFOEX))


                If GetVersionEx(osVersionInfo) Then
                    Dim majorVersion As Integer = osVersion.Version.Major
                    Dim minorVersion As Integer = osVersion.Version.Minor
                    Dim productType As Byte = osVersionInfo.wProductType
                    Dim suiteMask As Short = osVersionInfo.wSuiteMask


                    If majorVersion = 4 Then
                        If productType = VER_NT_WORKSTATION Then
                            tEdition = "Workstation"
                        ElseIf productType = VER_NT_SERVER Then
                            If (suiteMask And VER_SUITE_ENTERPRISE) <> 0 Then
                                tEdition = "Enterprise Server"
                            Else
                                tEdition = "Standard Server"
                            End If
                        End If


                    ElseIf majorVersion = 5 Then
                        If productType = VER_NT_WORKSTATION Then
                            If (suiteMask And VER_SUITE_PERSONAL) <> 0 Then
                                tEdition = "Home"
                            Else
                                tEdition = "Professional"
                            End If
                        ElseIf productType = VER_NT_SERVER Then
                            tEdition = "Server"
                        Else
                            If (suiteMask And VER_SUITE_ENTERPRISE) <> 0 Then
                                tEdition = "Enterprise"
                            Else
                                tEdition = "Standard"
                            End If
                        End If


                    ElseIf majorVersion = 6 Then
                        Dim ed As Integer
                        If GetProductInfo(majorVersion, minorVersion, osVersionInfo.wServicePackMajor, osVersionInfo.wServicePackMinor, ed) Then
                            Select Case ed
                                Case PRODUCT_BUSINESS
                                    tEdition = "Business"
                                    Exit Select
                                Case PRODUCT_DATACENTER_SERVER
                                    tEdition = "Datacenter Server"
                                    Exit Select
                                Case PRODUCT_ENTERPRISE
                                    tEdition = "Enterprise"
                                    Exit Select
                                Case PRODUCT_ENTERPRISE_SERVER
                                    tEdition = "Enterprise Server"
                                    Exit Select
                                Case PRODUCT_HOME_BASIC
                                    tEdition = "Home Basic"
                                    Exit Select
                                Case PRODUCT_HOME_PREMIUM
                                    tEdition = "Home Premium"
                                    Exit Select
                                Case PRODUCT_HOME_PREMIUM_SERVER
                                    tEdition = "Home Premium Server"
                                    Exit Select
                                Case PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT
                                    tEdition = "Windows Essential Business Management Server"
                                    Exit Select
                                Case PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING
                                    tEdition = "Windows Essential Business Messaging Server"
                                    Exit Select
                                Case PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY
                                    tEdition = "Windows Essential Business Security Server"
                                    Exit Select
                                Case PRODUCT_PROFESSIONAL
                                    tEdition = "Professional"
                                    Exit Select
                                Case PRODUCT_SERVER_FOUNDATION
                                    tEdition = "Server Foundation"
                                    Exit Select
                                Case PRODUCT_STANDARD_SERVER
                                    tEdition = "Standard Server"
                                    Exit Select
                                Case PRODUCT_STARTER
                                    tEdition = "Starter"
                                    Exit Select
                                Case PRODUCT_UNDEFINED
                                    tEdition = "Unknown product"
                                    Exit Select
                                Case PRODUCT_ULTIMATE
                                    tEdition = "Ultimate"
                                    Exit Select
                                Case PRODUCT_WEB_SERVER
                                    tEdition = "Web Server"
                                    Exit Select
                            End Select
                        End If
                    End If
                End If
                m_Edition = tEdition
                Return m_Edition
            End If
        End Get
    End Property


    Public ReadOnly Property Name() As String
        Get
            Dim osVersion As OperatingSystem = Environment.OSVersion
            Dim osVersionInfo As New OSVERSIONINFOEX()
            osVersionInfo.dwOSVersionInfoSize = Marshal.SizeOf(GetType(OSVERSIONINFOEX))


            If GetVersionEx(osVersionInfo) Then
                Dim majorVersion As Integer = osVersion.Version.Major
                Dim minorVersion As Integer = osVersion.Version.Minor


                Select Case osVersion.Platform
                    Case PlatformID.Win32S
                        m_Name = "Windows 3.1"
                        Exit Select
                    Case PlatformID.WinCE
                        m_Name = "Windows CE"
                        Exit Select
                    Case PlatformID.Win32Windows
                        If True Then
                            If majorVersion = 4 Then
                                Dim csdVersion As String = osVersionInfo.szCSDVersion
                                Select Case minorVersion
                                    Case 0
                                        If csdVersion = "B" OrElse csdVersion = "C" Then
                                            m_Name = "Windows 95 OSR2"
                                        Else
                                            m_Name = "Windows 95"
                                        End If
                                        Exit Select
                                    Case 10
                                        If csdVersion = "A" Then
                                            m_Name = "Windows 98 Second Edition"
                                        Else
                                            m_Name = "Windows 98"
                                        End If
                                        Exit Select
                                    Case 90
                                        m_Name = "Windows Me"
                                        Exit Select
                                End Select
                            End If
                            Exit Select
                        End If
                    Case PlatformID.Win32NT
                        If True Then
                            Dim productType As Byte = osVersionInfo.wProductType


                            Select Case majorVersion
                                Case 3
                                    m_Name = "Windows NT 3.51"
                                    Exit Select
                                Case 4
                                    Select Case productType
                                        Case 1
                                            m_Name = "Windows NT 4.0"
                                            Exit Select
                                        Case 3
                                            m_Name = "Windows NT 4.0 Server"
                                            Exit Select
                                    End Select
                                    Exit Select
                                Case 5
                                    Select Case minorVersion
                                        Case 0
                                            m_Name = "Windows 2000"
                                            Exit Select
                                        Case 1
                                            m_Name = "Windows XP"
                                            Exit Select
                                        Case 2
                                            m_Name = "Windows Server 2003"
                                            Exit Select
                                    End Select
                                    Exit Select
                                Case 6
                                    Select Case minorVersion
                                        Case 0
                                            Select Case productType
                                                Case 1
                                                    m_Name = "Windows Vista"
                                                    Exit Select
                                                Case 3
                                                    m_Name = "Windows Server 2008"
                                                    Exit Select
                                            End Select
                                            Exit Select


                                        Case 1
                                            Select Case productType
                                                Case 1
                                                    m_Name = "Windows 7"
                                                    Exit Select
                                                Case 3
                                                    m_Name = "Windows Server 2008 R2"
                                                    Exit Select
                                            End Select
                                            Exit Select
                                    End Select
                                    Exit Select
                            End Select
                            Exit Select
                        End If
                End Select
            End If


            Return m_Name
        End Get
    End Property


    Public ReadOnly Property ServicePack() As String
        Get
            Dim ssPack As String = [String].Empty
            Dim osVersionInfo As New OSVERSIONINFOEX()


            osVersionInfo.dwOSVersionInfoSize = Marshal.SizeOf(GetType(OSVERSIONINFOEX))


            If GetVersionEx(osVersionInfo) Then
                ssPack = osVersionInfo.szCSDVersion
            End If


            Return IIf(ssPack = "", "None", ssPack)
        End Get
    End Property
End Module
