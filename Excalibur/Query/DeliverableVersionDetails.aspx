<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Query_DeliverableVersionDetails" Codebehind="DeliverableVersionDetails.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
<% If ContentFormat = ContentFormats.HTML Then%>
    <title><%=Title%></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <link href="../style/excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../style/bubble.css" rel="stylesheet" type="text/css" />
    <link href="../style/deliverableversiondetails.css" rel="stylesheet" type="text/css" />

<script type="text/javascript">
<!--    script
    function updateLoadingMsg(s, n) {
        document.getElementById("loadingmsg" + n).innerHTML = s;
    }
-->
</script>
<%Else%>
    <style type="text/css">
        <%Response.WriteFile("../style/general.css")%>
        <%Response.WriteFile("../style/excalibur.css")%>
        <%Response.WriteFile("../style/bubble.css")%>
        <%Response.WriteFile("../style/deliverableversiondetails.css")%>
    </style>
<%End If%>
</head>
<body>
    <% If AdvancedSearch Then %>
        <% If ContentFormat = ContentFormats.HTML Then%>
            <div id="loading">Retrieving deliverables matching the specified search criteria. Please wait...<br />
            <span id="loadingmsg2"></span><span id="loadingmsg1"></span>
            <script type="text/javascript">
            <!--                script
                window.onload = function () {
                    document.getElementById("loading").style.display = "none";
                }
            -->
            </script>
            </div>
        <%End If%>
        <% 
            Response.Flush()
            GetDeliverables()
            If ReportDeliverablesCount > 0 And ReportAvailableDeliverablesCount > ReportDeliverablesCount Then %>
            <div id="pagemsg">
            NOTE: <%=ReportAvailableDeliverablesCount%> deliverables found matching the search criteria but only displaying
            maximum of <%=ReportDeliverablesCount%> deliverables. Please narrow your search criteria.
            </div>
            <%End If
                Response.Write("<div class=""title"">" & Title & "</div>")
    End If%>
    <div id="content">
    <% If ReportDeliverablesCount > 1 Then %>
        <a name="reportsummary"></a>
        <div id="reportsummary">
            # of Deliverables: <%=ReportDeliverablesCount%>
            <table class="summary">
            <tr><td class="tblHeader">Products</td><td class="tblHeader">ID</td><td class="tblHeader">Deliverables</td><td class="tblHeader">Vendor</td><td class="tblHeader">Developer</td><td class="tblHeader">Manager</td></tr>
            <%Do%>
                <tr>
                    <td><%=VP4Web.DOTSName%></td> <!-- display products -->
                    <% If IsValidQuery() Then%>
                        <td>
                             <%=VP4Web.VersionID%>
                        </td>
                        <td>
                            <% If ContentFormat = ContentFormats.HTML Then%>
                                <a href="#<%=VP4Web.VersionID%>"><%=VP4Web.FullName%></a>
                            <%Else%>
                                <%=VP4Web.FullName%>
                            <%End If%>
                        </td>
                        <td><%=VP4Web.VersionVendor%></td>
                        <td><%=VP4Web.Developer%></td>
                        <td><%=VP4Web.DevManager%></td>
                    <%Else%>
                        <td colspan="4"><%=VP4Web.ErrMsg%></td>
                    <%End If%>
                </tr>
            <%Loop While GetNextDeliverable()%>
            </table>
        </div>
        <%End If%>

    <%Do%>
    <div class="deliverable"><a name="<%=VP4Web.VersionID%>"></a>
    <% If IsValidQuery() Then %>
        <div class="subtitle"><%=VP4Web.FullName%>
        <% If AdvancedSearch And ReportDeliverablesCount > 1 And ContentFormat = ContentFormats.HTML Then%>
            <span class="totop">[<a href="#reportsummary">Back to Top</a>]</span>
        <%End If%>
        </div>
        <div class="details">
	        <table>

                <!-- Misc Info; id, version, etc. -->
		        <tr>
                <td><table>
                <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
                    <% If VP4Web.HFCN Then%>
		                <tr><td colspan="2">This is an HFCN release</td></tr>
                    <% End If%>
		            <tr><td class="label">ID:</td><td><%=VP4Web.VersionIDHTML%></td></tr>
		            <tr><td class="label">Hardware Version:</td><td><%=VP4Web.Version%></td></tr>
		            <tr><td class="label">Firmware Version:</td><td><%=VP4Web.Revision%></td></tr>
                    <tr><td class="label">Rev:</td><td><%=VP4Web.Pass%></td></tr>
                <% ElseIf VP4Web.TypeID = DeliverableTypes.Firmware Then%>
		            <tr><td class="label">ID:</td><td><%=VP4Web.VersionIDHTML%></td></tr>
		            <tr><td class="label">Version:</td><td><%=VP4Web.Version%></td></tr>
                    <% If VP4Web.HFCN And VP4Web.CategoryID = 161 Then%>
		                <tr><td colspan="2">This is an HFCN release</td></tr>
                    <% End If%>
                <% Else%>
		            <tr><td class="label">ID:</td><td><%=VP4Web.VersionIDHTML%></td></tr>
		            <tr><td class="label">Version:</td><td><%=VP4Web.Version%></td></tr>
		            <tr><td class="label">Revision:</td><td><%=VP4Web.Revision%></td></tr>
                    <tr><td class="label">Pass:</td><td><%=VP4Web.Pass%></td></tr>
                <% End If%>
		        </table></td>

	            <td><table>
	            <% If (VP4Web.SupplierID = 0) Or (VP4Web.SupplierID = VP4Web.VersionVendorID) Then%>
	                <tr><td class="label">Vendor/Supplier:</td><td><%=VP4Web.VersionVendor%></td></tr>
	            <% Else%>
                    <tr><td class="label">Vendor/Supplier:</td><td><%=VP4Web.VersionVendor & " (" & VP4Web.Supplier & ")"%></td></tr>
                <% End If%>
		        <tr><td class="label">Vendor Version:</td><td><%=VP4Web.VendorVersion%></td></tr>
                <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
		            <tr><td class="label">HP Part Number:</td><td><%=VP4Web.PartNumber%></td></tr>
		            <tr><td class="label">Model Number:</td><td><%=VP4Web.ModelNumber%></td></tr>
                <% Else%>
    		        <tr><td class="label">FileName:</td><td><%=VP4Web.Filename%></td></tr>
                <% End If%>
		        </table></td>

	            <td><table>
                <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
    		        <tr><td class="label">Code Name:</td><td><%=VP4Web.CodeName%></td></tr>
                <% End If%>
		        <tr><td class="label">Dev. Manager:</td><td><%=VP4Web.DevManager%></td></tr>
		        <tr><td class="label">Developer:</td><td><%=VP4Web.Developer%></td></tr>
                <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
                    <tr><td class="label">Production Level:</td><td><%=VP4Web.BuildLevel%></td></tr>
                <% Else%>
                    <tr><td class="label">Build Level:</td><td><%=VP4Web.BuildLevel%></td></tr>
                <% End If %>
                </table></td>
                </tr>

                <!-- Milestone -->
                <% If Not VP4Web.MilestoneList Is Nothing Then%>
                <tr><td colspan="3">
                    <div class="workflow">
                        <table>
                    	    <tr>
                                <td class="label">Workflow Step</td>
                                <td class="label">Status</td>
                                <td class="label">Planned Date</td>
                                <td class="label">Actual Date</td>
                             </tr>

                             <% For Each milestone As MilestoneRecord In VP4Web.MilestoneList %>
                                <tr>
                                    <td><%=milestone.Name%></td>
                                    <td><%=milestone.Status%></td>
                                    <td><%=milestone.PlannedDate%></td>
                                    <td><%=milestone.ActualDate%></td>
                                </tr>
                             <%Next%>
                        </table>
                    </div>
                </td></tr>
                <%End If%>

                <!-- Samples Available -->
                <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Samples Available Date:</span><%=VP4Web.SampleDate%>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="label">Confidence:</span>
                            <% If VP4Web.SampleConfidence = 1 Then%><span class="confidenceHigh">High
                            <% ElseIf VP4Web.SampleConfidence = 2 Then%><span class="confidenceMedium">Medium
                            <% ElseIf VP4Web.SampleConfidence = 3 Then%><span class="confidenceLow">Low
                            <% Else%><span class="confidenceUnknown">Unknown
                            <% End If %>
                            </span>
                        </td></tr>
                    </table></td></tr>
                <% End If %>

                <!-- Intro Date/Mass Production -->
                <% If VP4Web.IntroDate <> "" Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">
                            <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>Mass Production:
                            <% Else%>Intro Date:
                            <% End If%>
                            </span><%=VP4Web.IntroDate%>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="label">Confidence:</span>
                            <% If VP4Web.IntroConfidence = 1 Then%><span class="confidenceHigh">High
                            <% ElseIf VP4Web.IntroConfidence = 2 Then%><span class="confidenceMedium">Medium
                            <% ElseIf VP4Web.IntroConfidence = 3 Then%><span class="confidenceLow">Low
                            <% Else%><span class="confidenceUnknown">Unknown
                            <% End If %>
                            </span>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Available Until -->
                <% If VP4Web.EOLDate <> "" Or Not VP4Web.Active Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Available Until:</span>
                            <% If VP4Web.EOLDate <> "" Then%><%=VP4Web.EOLDate%>
                            <% Else%>Unknown
                            <% End If%>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <% If Not VP4Web.Active Then%>This version is End of Life<% End If %>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Special Notes -->
                <% If VP4Web.SpecialNotes <> "" And VP4Web.TypeID <> DeliverableTypes.Hardware And Left(VP4Web.Filename, 5) <> "HFCN_" Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Special Notes:</span><%=VP4Web.SpecialNotes%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Icons Installed -->
                <% If VP4Web.IconsInstalled <> "" And VP4Web.TypeID <> DeliverableTypes.Hardware Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Icons Installed:</span><%=VP4Web.IconsInstalled%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Property Tabs -->
                <% If VP4Web.PropertyTabs <> "" And VP4Web.TypeID <> DeliverableTypes.Hardware Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Property Tabs Added:</span><%=VP4Web.PropertyTabs%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Packaging -->
                <% If VP4Web.Packaging <> "" And VP4Web.TypeID <> DeliverableTypes.Hardware And _
                       (VP4Web.TypeID <> DeliverableTypes.Firmware Or VP4Web.Preinstall Or VP4Web.FloppyDisk Or VP4Web.CDImage Or _
                        VP4Web.Scriptpaq) Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Packaging:</span><%=VP4Web.Packaging%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- ROM Components -->
                <% If VP4Web.ROMComponents <> "" And VP4Web.TypeID <> DeliverableTypes.Hardware And _
                       (VP4Web.TypeID = DeliverableTypes.Firmware Or VP4Web.Binary Or VP4Web.Rompaq Or VP4Web.PreinstallROM Or _
                        VP4Web.CAB) Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">ROM Components:</span><%=VP4Web.ROMComponents%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Replicated By -->
                <% If VP4Web.ISOImage Or VP4Web.AR Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Replicated By:</span><%=VP4Web.Replicater%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- CD/DVD Part Number / Kit Number -->
                <% If Left(VP4Web.Filename, 5) <> "HFCN_" And (VP4Web.CDImage Or VP4Web.ISOImage) Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">CD/DVD Part Number:</span>
                            <% If VP4Web.Multilangauge Then%>
                                <%=VP4Web.CDPartNumber%>
                                </td><td>
                                <span class="label">Kit Number:</span>
                                <%=VP4Web.CDKitNumber%>
                            <% Else %>
                                <%=VP4Web.PartNumbers%>
                            <% End If%>
                        </td></tr>
                    </table></td></tr>

                    <% If Not VP4Web.Multilangauge Then%>
                        <tr><td colspan="3"><table>
                            <tr><td>
                                <span class="label">Kit Number:</span>
                                <%=VP4Web.CDKitNumbers%>
                            </td></tr>
                        </table></td></tr>
                    <% End If %>
                <% End If%>

                <!-- RoHS/Green Spec -->
                <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">RoHS/Green Spec:</span>
                            <%=VP4Web.RoHSGreenSpec%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Functional Spec -->
                <% If VP4Web.DeliverableSpec <> "" Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Functional Spec:</span>
                            <%=VP4Web.DeliverableSpec%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Location/Path -->
                <% If Not VP4Web.AR Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Location/Path:</span>
                            <%=VP4Web.ImagePath%>
                        </td></tr>
                    </table></td></tr>
                <% End If%>

                <!-- Comments -->
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">Comments:</span>
                        <%=VP4Web.Comments%>
                    </td></tr>
                </table></td></tr>

                <!-- OTS List -->
                <% If Not VP4Web.OTSList Is Nothing Then%>
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">Observations fixed in this release:</span>
                        <% For Each ots As OTSRecord In VP4Web.OTSList %>
                            <%=ots.HTMLLink%><br />
                        <%Next%>
                    </td></tr>
                </table></td></tr>
                <%End If%>

                <!-- Changes -->
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">Modifications, Enhancements, or Reason for Release:</span>
                        <br /><%=VP4Web.Changes%>
                    </td></tr>
                </table></td></tr>

                <!-- Supported Operating Systems -->
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">Supported Operating Systems:</span>
                        <% If VP4Web.SelectedOSes <> "" Then%><%=VP4Web.SelectedOSes%>
                        <% Else%>OS Independent
                        <% End If%>
                    </td></tr>
                </table></td></tr>

                <!-- Supported Languages/Countries -->
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">
                        <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>Supported Countries:
                        <% Else%>Supported Languages:
                        <% End If%>
                        </span>
                        <% If VP4Web.SelectedLangs <> "" Then%><%=VP4Web.SelectedLangs%>
                        <% ElseIf VP4Web.TypeID = DeliverableTypes.Hardware Then%>Country Independent
                        <% Else%>Language Independent
                        <% End If%>
                    </td></tr>
                </table></td></tr>

                <!-- PNP Devices dependencies -->
                <%If Not VP4Web.PNPDevices Is Nothing Then%>
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">PNP Devices dependencies:</span>
                        <% For Each pnpDev As String In VP4Web.PNPDevices%>
                            <%=pnpDev%><br />
                        <%Next%>
                    </td></tr>
                </table></td></tr>
                <%End If%>

                <!-- Deliverable dependencies -->
                <% If VP4Web.DeliverableDependencies <> "" Then%>
                    <tr><td colspan="3"><table>
                        <tr><td>
                            <span class="label">Deliverables Dependencies:</span>
                            <%=VP4Web.DeliverableDependencies%><br />
                        </td></tr>
                    </table></td></tr>
                <% End If %>

                <!-- Other dependencies -->
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">Other Dependencies:</span>
                        <%=VP4Web.SWDependencies%><br />
                    </td></tr>
                </table></td></tr>

                <!-- Products -->
                <%If Not VP4Web.ProductList Is Nothing Then%>
                <tr><td colspan="3"><table>
                    <tr><td>
                        <span class="label">Products:</span>
                        </td>
                        <td class="fullwidth">
                            <% If UBound(VP4Web.ProductList) < 1 Then%>No Product Status Available
                            <% Else%>
                                <div class="products">
                                    <table>
                                        <% If VP4Web.TypeID = DeliverableTypes.Hardware Then%>
                                            <tr class="tblHeader">
                                                <td class="tblHeader">Product</td><td class="tblHeader">Hardware PM</td>
                                                <td class="tblHeader">Dev. Approval</td><td class="tblHeader">Status</td>
                                                <% If VP4Web.ProductListHasTestingSummary Then%>
                                                <td class="tblHeader">Testing Summary</td>
                                                <% End If  %>
                                                <td class="tblHeader">Notes</td>
                                                <td class="tblHeader">Pilot Status</td><td class="tblHeader">Restrictions</td>
                                            </tr>
                                            <% For Each p As ProductRecord In VP4Web.ProductList%>
                                                <tr>
                                                    <td class="nowrap"><%=p.Name%></td><td class="nowrap"><%=p.ProjectManager%></td>
                                                    <td class="nowrap"><%=p.DeveloperStatus%></td><td class="nowrap"><%=p.Status%></td>
                                                    <% If VP4Web.ProductListHasTestingSummary Then%>
                                                    <td><%=p.TestingSummary%></td>
                                                    <% End If  %>
                                                    <td><%=p.PINOrTestNotes%></td>
                                                    <td><%=p.PilotStatus%></td><td><%=p.Restrictions%></td>
                                                </tr>
                                            <% Next %>
                                        <% Else%>
                                            <tr class="tblHeader">
                                                <td class="tblHeader">Product</td><td class="tblHeader">SE PM</td>
                                                <td class="tblHeader">Dev. Approval</td><td class="tblHeader">SE PM Status</td>
                                                <td class="tblHeader">Preinstall</td>
                                            </tr>
                                            <% For Each p As ProductRecord In VP4Web.ProductList%>
                                                <tr>
                                                    <td class="nowrap"><%=p.Name%></td><td class="nowrap"><%=p.ProjectManager%></td>
                                                    <td class="nowrap"><%=p.DeveloperStatus%></td><td class="nowrap"><%=p.Status%></td>
                                                    <td><%=p.PINOrTestNotes%></td>
                                                </tr>
                                            <% Next %>
                                        <% End If %>
                                    </table>
                                </div>
                            <% End If %>
                    </td></tr>
                </table></td></tr>
                <%End If%>

            </table>
        </div>
    <% Else%>
        <%=ErrorMsg%>
    <% End If%>
    </div>
    <%Loop While GetNextDeliverable()%>

    </div>

    <div class="confidential">HP Restricted</div>
</body>
</html>
