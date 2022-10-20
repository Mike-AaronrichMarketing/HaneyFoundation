Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.DateTime
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.IO
Imports System.Net
Imports System.Net.Mail

Partial Class ADMIN_Contacts
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then

			txtCurDate.Text = Now()
			txtFormName.Text = "Planned Giving"

			If Len(Server.HtmlDecode(Request.QueryString("mNum"))) > 0 Then

				Dim mNum As Integer = Request.QueryString("mNum")
				'Dim currentStatus As String = ""
				'Dim fileName As String

				lblmNum.Text = mNum.ToString
		
				Dim oItem As New PC_Contacts
		
				oItem = DAL_Contacts.GetContactsByNum(mNum)
		
				With oItem
					txtFormName.Text = .FormName
					txtfirstName.Text = .firstName
					txtlastName.Text = .lastName
					txtcompanyName.Text = .companyName
					txtaddress1.Text = .address1
					txtaddress2.Text = .address2
					txtcity1.Text = .city1
					ddlstate1.SelectedIndex = ddlstate1.Items.IndexOf(ddlstate1.Items.FindByValue(.state1))
					txtzipCode.Text = .zipCode
					txtphoneNumber.Text = .phoneNumber
					txtemailAddress.Text = .emailAddress
					'txtdateOfBirth.Text = .dateOfBirth
					txtcomments.Text = .comments
					chkmailingList.Text = .mailingList
					ddlamount.SelectedIndex = ddlamount.Items.IndexOf(ddlamount.Items.FindByValue(.amount))
					'ddlfrequency.SelectedIndex = ddlfrequency.Items.IndexOf(ddlfrequency.Items.FindByValue(.frequency))
					txtCurDate.Text = Now()
				End With

				'Submit.Text = "Update"

			End If

        End If

    End Sub

	Protected Sub submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submit.Click

		Dim oItem As New PC_Contacts
		Dim Rslt As Integer
		Dim cnt As Integer
		'Dim fileName As String
		txtFormName.Text = "Planned Giving"
		lblFormName.Text = "Planned Giving"

		If Len(txtURL.text) > 1 Then
			Response.Redirect("Error.aspx")
		End If

		With oItem
			.FormName = txtFormName.Text
			.firstName = txtfirstName.Text
			.lastName = txtlastName.Text
			.companyName = txtcompanyName.Text
			.address1 = txtaddress1.Text
			.address2 = txtaddress2.Text
			.city1 = txtcity1.Text
			.state1 = ddlstate1.SelectedValue()
			.zipCode = txtzipCode.Text
			.phoneNumber = txtphoneNumber.Text
			.emailAddress = txtemailAddress.Text
			'.dateOfBirth = txtdateOfBirth.Text
			.comments = txtcomments.Text
			.mailingList = replace(replace(chkmailingList.checked, "True", "Yes"), "False", "No")
			.amount = ddlamount.SelectedValue()
			'.frequency = ddlfrequency.SelectedValue()
			.CurDate = Now()
			'If SaveButton.Text = "Update Item" Then
			'.mNum = lblmNum.Text
			'.mNum = convert.toInt16(lblmNum.Text)
			'End If
		End With

		'If SaveButton.Text = "Update Item" Then

		'Rslt = DAL_Contacts.ModContacts(oItem)

		'Else

		Rslt = DAL_Contacts.AddPlannedGivingContacts(oItem)

		With oItem
			.FormName = txtFormName.Text
			.firstName = txtfirstName.Text
			.lastName = txtlastName.Text
			.companyName = txtcompanyName.Text
			.address1 = txtaddress1.Text
			.address2 = txtaddress2.Text
			.city1 = txtcity1.Text
			.state1 = ddlstate1.SelectedValue()
			.zipCode = txtzipCode.Text
			.phoneNumber = txtphoneNumber.Text
			.emailAddress = txtemailAddress.Text
			'.dateOfBirth = txtdateOfBirth.Text
			.comments = txtcomments.Text
			.mailingList = replace(replace(chkmailingList.checked, "True", "Yes"), "False", "No")
			.amount = ddlamount.SelectedValue()
			'.frequency = ddlfrequency.SelectedValue()
			.CurDate = Now()
			'If SaveButton.Text = "Update Item" Then
			'	.mNum = lblmNum.Text
			'End If
		End With
		'Submit.Text = "Submit" 

		Dim rValidate As Boolean = IsGoogleCaptchaValid()
		If rValidate = True Then
			Dim mail As New MailMessage()

			mail.From = New MailAddress("DONOTREPLY@haneyfoundation.org")
			'mail.To.Add("" & txtemailAddress.Text & "")
			mail.To.Add("info@haneyfoundation.org")
			'mail.CC.Add("marketing@adoptme.org")
			mail.Bcc.Add("armwebforms@aaronrich.com")

			mail.Subject = "Haney Foundation :: Planned Giving"
			mail.IsBodyHtml = True
			'mail.Body = "Thank you for your donation to the Humane Society of Bay County!  We could not provide an impact in the animal community without the generosity of our donors.  If you need any additional information please feel free to contact us directly. <br><br>" & _
			mail.Body = "<strong>Name: </strong>" & txtfirstName.Text & " " & txtlastName.Text & "<br>" &
			"<strong>Company Name: </strong>" & txtcompanyName.Text & "<br>" &
			"<strong>Address: </strong>" & txtaddress1.Text & " " & txtaddress2.Text & ", " & txtcity1.Text & ", " & ddlState1.Text & " " & txtzipCode.Text & "<br>" &
			"<strong>Phone Number: </strong>" & txtphoneNumber.Text & "<br>" &
			"<strong>Email Address: </strong>" & txtemailAddress.Text & "<br>" &
			"<strong>Mailing List: </strong>" & Replace(Replace(chkmailingList.Checked, "True", "Yes"), "False", "No") & "<br>" &
			"<strong>Planned Giving Amount: </strong>" & ddlamount.Text & "<br>" &
			"<strong>Comments: </strong>" & txtcomments.Text

			Dim smtp As New SmtpClient()

			Try
				smtp.Credentials = New _
			Net.NetworkCredential("PM-T-outbound-9dZRbjUtzy3YL6n2wHmNHZ", "Ja2s7l5trnQW_ZgQao6tCX-6fLpZIZoqcEoE")
				smtp.Port = 587
				smtp.Host = "smtp.postmarkapp.com"

				smtp.Send(mail)
				Response.Redirect("Thanks.aspx")
			Catch ehttp As System.Web.HttpException
				Response.Redirect("Error.aspx")
			End Try






			If ddlamount.text = "$5" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=U5WM3NHWBSMQG")
			ElseIf ddlamount.text = "$10" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=4S6MS75AWYLVW")
			ElseIf ddlamount.text = "$20" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=BBM6VLY3GHWPG")
			ElseIf ddlamount.text = "$50" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=3QE998YMQM6QA")
			ElseIf ddlamount.text = "$100" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=XJYTNA59L8QX4")
			ElseIf ddlamount.text = "$500" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=74E766VTKBFRQ")
			ElseIf ddlamount.text = "$1,000" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=NV9YVG95LG8LQ")
			ElseIf ddlamount.text = "Other" Then
				Response.Redirect("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=MS8VEZCRYB458")

			ElseIf ddlamount.text = "" Then
				Response.Redirect("Error.aspx")
			Else Response.Redirect("Thanks.aspx")

			End If
		

		End If

	End Sub
	Private Function IsGoogleCaptchaValid() As Boolean
		Dim recaptchaResponse As String = Request.Form("g-recaptcha-response")
		Try
			If Not String.IsNullOrEmpty(recaptchaResponse) Then
				Dim request As WebRequest = WebRequest.Create("https://www.google.com/recaptcha/api/siteverify?secret=6Lf4oRkiAAAAAOKIC42IHNKSdyT_YnbodGsnzcNg&response=" + recaptchaResponse)
				request.Method = "POST"
				request.ContentType = "application/json; charset=utf-8"
				Dim postData As String = ""

				'get a reference to the request-stream, and write the postData to it
				Using s As IO.Stream = request.GetRequestStream()
					Using sw As New IO.StreamWriter(s)
						sw.Write(postData)
					End Using
				End Using
				''get response-stream, and use a streamReader to read the content
				Using s As IO.Stream = request.GetResponse().GetResponseStream()
					Using sr As New IO.StreamReader(s)
						'decode jsonData with javascript serializer
						Dim jsonData = sr.ReadToEnd()
						'message.value += jsonData
						If jsonData.Contains("""success"": true") = True Then
							Return True
						End If
					End Using
				End Using
			End If
		Catch ex As Exception
			txtcomments.text = ex.ToString()
		End Try

		'message.value = recaptchaResponse
		Return False
	End Function

End Class
