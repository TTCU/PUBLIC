# Microsoft Teams Call History Outlook/OWA Addin #
![](https://1.bp.blogspot.com/-WJGIYKDZ2QM/YJjCw-AMckI/AAAAAAAAOGo/HcTdu6C4qH497My59jGYaJYe-xWpC2zoACLcBGAsYHQ/w640-h530/callhist1.PNG)


```
<soap:Envelope
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
	xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
	xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
	<soap:Header>
		<t:RequestServerVersion Version="Exchange2016" />
	</soap:Header>
	<soap:Body>
		<m:FindItem Traversal="Shallow">
			<m:ItemShape>
				<t:BaseShape>IdOnly</t:BaseShape>
				<t:AdditionalProperties>
					<t:FieldURI FieldURI="item:ItemClass"
						xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" />
						<t:ExtendedFieldURI PropertyTag="96" PropertyType="SystemTime" />
						<t:ExtendedFieldURI PropertyTag="97" PropertyType="SystemTime" />
						<t:ExtendedFieldURI PropertyTag="23809" PropertyType="String" />
					</t:AdditionalProperties>
				</m:ItemShape>
				<m:IndexedPageItemView MaxEntriesReturned="40" Offset="0" BasePoint="Beginning" />
				<m:ParentFolderIds>
					<t:FolderId Id="A=....." />
				</m:ParentFolderIds>
				<m:QueryString>participants:"e5tmp5@datarumble.com"</m:QueryString>
			</m:FindItem>
		</soap:Body>
	</soap:Envelope>
```

The extended properties that are used for the Start and End are pidTagStart and pidTagEnd which get set on the CDR mailbox item and then the SenderEmailSMTP address is used to work out who is the Organizer or the call initiator. To work out the type of Call the ItemClass is used so the last part of the ItemClass will tell you what the CRD is for eg call or meeting (If your wondering if you could do the same with the Microsoft Graph the answer is no because the Items in question are not a subclass of IPM.Note so aren't accessible when using the Microsoft Graph)

!![](https://1.bp.blogspot.com/-fV2Wxo0Cr7Q/XIHvqow9SVI/AAAAAAAACTM/Vo-Pc3Q74AgZ-_KNh_Nl9UhZmBbouJS1wCLcBGAs/s1600/addin3.JPG)
