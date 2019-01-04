Imports System.IO


Public Class Form1
    Dim table As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Dim info As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Dim addresses As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Dim links As New List(Of TextBox)
    Dim user As String = Environment.UserName
    Dim fileName As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        createTable()
        web.Text = "priorityfinancial.net"
        channels.SelectedIndex = 0
    End Sub

    Private Sub Generate_Click(sender As Object, e As EventArgs) Handles Generate.Click
        Dim list As New List(Of TextBox)
        Dim counter As Integer = 0
        Dim writeLine As Boolean = True
        links.Clear()
        Generate.Enabled = False

        addresses.Add("address", address.Text)
        addresses.Add("suite", "Ste. " & suite.Text)
        addresses.Add("city", city.Text)
        addresses.Add("state", state.Text & " ")
        addresses.Add("zip", zip.Text)

        list.Add(txtName)
        list.Add(title)
        list.Add(email)
        list.Add(phone)
        list.Add(directphone)
        list.Add(cellphone)
        list.Add(faxphone)
        list.Add(nmls)
        list.Add(dre)

        If sigName.Text = "" Then
            fileName = "defaultSignature"
        Else
            fileName = sigName.Text
        End If

        For Each item As TextBox In list
            info.Add(item.Name, item.Text)
        Next
        For Each val As Object In pnlSocial.Controls
            Try
                links.Add(val)
            Catch c As InvalidCastException
                'do nothing
            End Try
        Next
        links.Remove(secure)
        If txtName.Text = "" Then
            txtName.BackColor = Color.LightCoral
            Return
        End If

        Try
            My.Computer.FileSystem.DeleteFile("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm")
        Catch
        End Try
        For Each current As String In File.ReadLines("htmlDoc.txt")
            If current.Contains("NAME") Then
                current = current.Replace("NAME", info.Item("txtName"))
            End If
            If current.Contains("TITLE") Then
                current = current.Replace("TITLE", "")
                If info.Item("title") <> "" Then
                    addTitle()
                End If
            End If
            If current.Contains("PHONE") Then
                current = current.Replace("PHONE", "")
                addPhones()
            End If
            If current.Contains("EMAIL/WEB") Then
                current = current.Replace("EMAIL/WEB", "")
                addEmail()
            End If
            If current.Contains("NMLSID") Then
                current = current.Replace("NMLSID", "")
                addNMLS()
            End If
            If current.Contains("CENTER") Then
                current = current.Replace("CENTER", "")
                If webcenter.Text <> "" Then
                    addWebcenter()
                End If
                addSocialMedia()
            End If
            If current.Contains("ADDRESS") Then
                current = current.Replace("ADDRESS", "")
                addAddress()
            End If
            If current.Contains("SECUREUPLOAD") Then
                current = ""
                If secure.Text <> "" Then
                    addSecure()
                End If
            End If
            If writeLine = True Then
                My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", current, True)
                My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                                "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", Environment.NewLine, True)
            End If
        Next
        Dim di As DirectoryInfo = New DirectoryInfo("C:\Users\" & user &
                           "\AppData\Roaming\Microsoft\Signatures\" & fileName & "_files")
        di.Create()

        MessageBox.Show("New Signature added as '" & fileName & "'")
        Me.Close()

    End Sub

    Sub addAddress()
        Dim flag As Boolean = False
        For Each val As KeyValuePair(Of String, String) In addresses
            If val.Value <> "" Then
                flag = True
            End If
        Next

        If flag = True Then
            My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", "<tr><td>", True)
            For Each val As KeyValuePair(Of String, String) In addresses
                If val.Value <> "" Then
                    Dim element As XElement
                    If val.Key = "state" Or val.Key = "zip" Then
                        element = <font style="color: rgb(88,88,90); font-size: 10pt; font-family: Poppins"><%= val.Value %></font>
                    Else
                        element = <font style="color: rgb(88,88,90); font-size: 10pt; font-family: Poppins"><%= val.Value %>, </font>
                    End If
                    My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                                "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)
                End If
            Next
            My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", "</td></tr>", True)
        End If
    End Sub

    Sub addSecure()
        Dim element As XElement
        element = <tr>
                      <td style="max-width: 500px;" colspan="2"><a href=<%= secure.Text %>><img style="width: 100%;" src="https://images.pfnrates.com/email_disclaimer.png"></img></a></td>
                  </tr>
        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)
    End Sub

    Sub addWebcenter()
        Dim element As XElement = <td colspan="5" data-bind="visible: bannerImageUrl() !== '', html: promoBannerOptionalLink">
                                      <a href=<%= webcenter.Text %>>
                                          <img id="TemplateBanner" data-class="external" src="https://images.pfnrates.com/email_apply_now.png" alt="applyNow" style="display: block"></img>
                                      </a>
                                  </td>
        Dim temp As String = element.ToString
        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", temp, True)
    End Sub

    Sub addEmail()
        Dim element As XElement
        Dim element2 As XElement
        Dim flag As Boolean = False
        Dim secondary As Boolean = False
        If info.Item("email") <> "" And web.Text <> "" Then
            flag = True
            If info.Item("email").Length + web.Text.Length > 43 Then
                secondary = True
                element = <tr>
                              <td style="font-family: Poppins; font-size: 10pt;"><font style="color: #333333; font-size: 10pt; font-family: Poppins">E: <a href="mailto:" style="color: #06c"><%= info.Item("email") %></a></font></td>
                          </tr>
                element2 = <tr>
                               <td style="font-family: Poppins; font-size: 10pt;"><font style="color: #333333; font-size: 10pt; font-family: Poppins">W: <a href=<%= web.Text %> style="color: #06c"><%= web.Text %></a></font></td>
                           </tr>
            Else
                element = <tr>
                              <td style="font-family: Poppins; font-size: 10pt;"><font style="color: #333333; font-size: 10pt; font-family: Poppins">E: <a href="mailto:" style="color: #06c"><%= info.Item("email") %></a></font><font style="color: #333333; font-size: 10pt; font-family: Poppins"> | </font><font style="color: #333333; font-size: 10pt; font-family: Poppins">W: <a href=<%= web.Text %> style="color: #06c"><%= web.Text %></a></font></td>
                          </tr>
            End If
        End If
        If info.Item("email") <> "" And web.Text = "" Then
            flag = True
            element = <tr>
                          <td style="font-family: Poppins; font-size: 10pt;"><font style="color: #333333; font-size: 10pt; font-family: Poppins">E: <a href="mailto:" style="color: #06c"><%= info.Item("email") %></a></font></td>
                      </tr>
        End If
        If info.Item("email") = "" And web.Text <> "" Then
            flag = True
            element = <tr>
                          <td style="font-family: Poppins; font-size: 10pt;"><font style="color: #333333; font-size: 10pt; font-family: Poppins">W: <a href=<%= web.Text %> style="color: #06c"><%= web.Text %></a></font></td>
                      </tr>
        End If
        If flag = True Then
            Dim temp As String = element.ToString
            temp.Replace("mailto:", "mailto:" & info.Item("email"))
            My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", temp, True)
        End If
        If secondary = True Then
            My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element2.ToString, True)
        End If
    End Sub

    Sub addNMLS()
        Dim element As New XElement("element")
        If info.Item("nmls") <> "" And info.Item("dre") <> "" Then
            element = <tr>
                          <td>
                              <font style="color: rgb(136, 136, 136); font-size: 10pt; font-family: Poppins">NMLS: <%= info.Item("nmls") %> | </font><font style="color: rgb(136, 136, 136); font-size: 10pt; font-family: Poppins"><%= licType.Text %>: <%= info.Item("dre") %></font>
                          </td>
                      </tr>
        End If
        If info.Item("nmls") <> "" And info.Item("dre") = "" Then
            element = <tr>
                          <td>
                              <font style="color: rgb(136, 136, 136); font-size: 10pt; font-family: Poppins">NMLS: <%= info.Item("nmls") %></font>
                          </td>
                      </tr>
        End If
        If info.Item("nmls") = "" And info.Item("dre") <> "" Then
            element = <tr>
                          <td>
                              <font style="color: rgb(136, 136, 136); font-size: 10pt; font-family: Poppins"><%= licType.Text %>: <%= info.Item("dre") %></font>
                          </td>
                      </tr>
        End If
        If element.IsEmpty Then
        Else
            My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)
        End If
    End Sub
    Sub openRow()
        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                                "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", "</td><tr>", True)
    End Sub
    Sub closeRow()
        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                                "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", "<tr><td style="" font-family:Poppins; font-size: 10pt;"">", True)
    End Sub

    Sub addPhones()
        Dim oneTime As Boolean = True
        Dim counter As Integer = 0
        Dim incrementer As Integer = 0
        Dim max As Integer
        For Each pair As KeyValuePair(Of String, String) In info
            If pair.Key.Contains("phone") And pair.Value <> "" Then
                counter += 1
            End If
        Next
        max = counter
        If counter > 0 Then
            closeRow()
            For Each pair As KeyValuePair(Of String, String) In info
                If pair.Key.Contains("phone") And pair.Value <> "" Then
                    Dim element As XElement
                    If max = 4 Then
                        If counter = 1 Or counter = 3 Then
                            element = <font style="color: black; font-size: 10pt; font-family: Poppins"><%= pair.Key.Substring(0, 1).ToUpper %>: <a href="tel:PHONE" style="color: #06c"><%= pair.Value %></a></font>
                        Else
                            element = <font style="color: black; font-size: 10pt; font-family: Poppins"><%= pair.Key.Substring(0, 1).ToUpper %>: <a href="tel:PHONE" style="color: #06c"><%= pair.Value %></a> | </font>
                        End If
                        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                                "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)
                        counter -= 1
                        If counter = 2 Then
                            openRow()
                            closeRow()
                        End If
                    Else
                        If counter = 1 Then
                            element = <font style="color: black; font-size: 10pt; font-family: Poppins"><%= pair.Key.Substring(0, 1).ToUpper %>: <a href="tel:PHONE" style="color: #06c"><%= pair.Value %></a></font>
                        Else
                            element = <font style="color: black; font-size: 10pt; font-family: Poppins"><%= pair.Key.Substring(0, 1).ToUpper %>: <a href="tel:PHONE" style="color: #06c"><%= pair.Value %></a> | </font>
                        End If
                        counter -= 1
                        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                                "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)
                    End If


                End If
            Next
            openRow()
        End If
    End Sub

    Sub addTitle()
        Dim current As String = info.Item("title")
        Dim element As XElement = <tr>
                                      <td style="padding-bottom: 5px; font-weight: bold; font-family: Poppins; font-size: 10pt; color: rgb(51, 51, 51);">
                                          <font style="font-family: Poppins; font-size: 10pt; color: rgb(51, 51, 51);"><%= current %></font>
                                      </td>
                                  </tr>
        My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)

    End Sub

    Sub addSocialMedia()
        Dim i As Integer
        For i = 0 To links.Count - 3
            If links(i).Text <> "" Then
                Dim current As String = table.Item(links(i).Name)
                Dim element As XElement = <td><a href=<%= links(i).Text %>><img data-class="external" style="display: inline" src=<%= current %>></img></a></td>
                My.Computer.FileSystem.WriteAllText("C:\Users\" & user &
                                            "\AppData\Roaming\Microsoft\Signatures\" & fileName & ".htm", element.ToString, True)
            End If
        Next
    End Sub

    Sub createTable()
        table.Add("facebook", "https://images.pfnrates.com/email_facebook.png")
        table.Add("twitter", "https://images.pfnrates.com/email_twitter.png")
        table.Add("linkedIn", "https://images.pfnrates.com/email_linkedin.png")
        table.Add("yelp", "https://images.pfnrates.com/email_yelp.png")
        table.Add("instagram", "https://images.pfnrates.com/email_insta.png")
        table.Add("youtube", "https://68ef2f69c7787d4078ac-7864ae55ba174c40683f10ab811d9167.ssl.cf1.rackcdn.com/youtube-icon_square_32x32.png")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        For Each control As Control In pnlAddress.Controls
            If TypeOf (control) Is TextBox Then
                control.Text = ""
            End If
        Next
    End Sub
End Class
