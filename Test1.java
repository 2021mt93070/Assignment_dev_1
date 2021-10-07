package com.arthrex.px.amcRedlineReport;

/** 
 * File                 : AMCReportDeletionOnSaveAs.java
 * Version   Date        Developer          Description
 * ------- ---------    -------------	   ---------------
 *  1.0	   03-Feb-21  Vijaya Gummapu	   Initial Version as part of Feb 2021 Release	 	
 */

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.LoggerContext;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IWFChangeStatusEventInfo;
import com.arthrex.common.Common;
import com.arthrex.px.amcRedlineReport.constants.AMCRedlineReportConstants;
import com.arthrex.px.amcRedlineReport.service.AMCRedlineReportService;
import com.arthrex.px.amcRedlineReport.service.CreateAMCExcelReport;
import com.arthrex.px.amcRedlineReport.util.AMCReportUtil;

/**
 * This class implements the methods for generating the AMC Redline Report on
 * AMC from pending to Approval status and attaching it to its Attachments tab.
 */

public class AMCRedlineReport implements IEventAction {

	private static final Logger pxLogger = LogManager.getLogger(AMCRedlineReport.class);
	AMCRedlineReportConstants constObj = AMCRedlineReportConstants.getInstance();
	LoggerContext context = (LoggerContext) LogManager.getContext(false);
	String formattedString = "";

	boolean newFlag = false;
	boolean revisedFlag = false;
	File file = null;
	List<File> repFile = new ArrayList<>();
	ArrayList<IItem> validItemList = new ArrayList<>();
	List<String> changeDetailsList = new ArrayList<>();

	Map<String, List<String[]>> partDetailsMap = new HashMap<>();
	List<String[]> newPartList = new ArrayList<>();
	List<String[]> revPartList = new ArrayList<>();

	Map<String, List<String[]>> rawMatDetMap = new HashMap<>();
	List<String[]> newRMList = new ArrayList<>();
	List<String[]> revRMList = new ArrayList<>();

	Map<String, List<String[]>> stdKeyDetMap = new HashMap<>();
	List<String[]> newStdList = new ArrayList<>();
	List<String[]> revStdList = new ArrayList<>();

	Map<String, String[]> itemDetailsMap = new HashMap<>();
	SimpleDateFormat simpleDateFormat = null;

	Map<String, List<String[]>> documentMap = new HashMap<>();
	List<String[]> newDocList = new ArrayList<>();
	List<String[]> revDocList = new ArrayList<>();

	Map<String, List<String[]>> refOppSheetMap = new HashMap<>();
	List<String[]> refOppSheetRevisedList = new ArrayList<>();
	List<String[]> refOppSheetNewList = new ArrayList<>();

	Map<String, List<String[]>> bomSheetMap = new HashMap<>();
	List<String[]> bomSheetRevisedList = new ArrayList<>();
	List<String[]> bomSheetNewList = new ArrayList<>();

	Map<String, List<String[]>> routingsheetMap = new HashMap<>();
	List<String[]> routingSheetHeaderNewList = new ArrayList<>();
	List<String[]> routingSheetHeaderRevisedList = new ArrayList<>();

	Map<String, List<String[]>> pvSheetMap = new HashMap<>();
	List<String[]> pvSheetRevisedList = new ArrayList<>();
	List<String[]> pvSheetNewList = new ArrayList<>();

	IAgileSession adminSession;

	String pxName = "PX AMCReportDeletionOnSaveAs Ver ";
	String strMessage = "[FAILED]: " + pxName + " failed due to an error. ";

	AMCReportUtil repUtil = new AMCReportUtil();
	AMCRedlineReportService service = new AMCRedlineReportService();

	public EventActionResult doAction(IAgileSession session, INode node, IEventInfo request) {

		IWFChangeStatusEventInfo eventInfo = (IWFChangeStatusEventInfo) request;

		formattedString = String.format("Logger file Location > %s", constObj.amcredlinelog4jpath);
		pxLogger.info(formattedString);
		file = new File(constObj.amcredlinelog4jpath);
		context.setConfigLocation(file.toURI());
		String strType = "";

		String strLabel = constObj.amcLabel;
		String strAMIDMS = constObj.amcAMIDMS;
		String strDrawing = constObj.amcDrawing;
		String strLabelDoc = constObj.amcLabelDocument;
		String strLabelFormat = constObj.amcLabelFormat;
		String strManual = constObj.amcManual;
		String strMfgDoc = constObj.amcManufacturingDocument;
		String strMfgInsp = constObj.amcManufacturingInspection;
		String strRawMaterial = constObj.amcRawMaterial;
		String strFGType = constObj.amcFinishedGood;
		String strComp = constObj.amcComponent;
		String strPackMaterail = constObj.amcPackagingMaterial;
		String strDirForUse = constObj.amcDirectionsForUse;
		String strMfgStdTextKey = constObj.amcMfgRoutingStdTextKey;
		String strNewRev = "";

		ActionResult actionResult = null;
		IChange amcObj = null;

		try {

			Common.initConfigurator();
			this.context.setConfigLocation(this.file.toURI());

			pxLogger.info(" ################# Redline Report Begins ################ ");

			AMCRedlineReportService.setLogger(pxLogger);
			AMCReportUtil.setLogger(pxLogger);
			CreateAMCExcelReport.setLogger(pxLogger);

			amcObj = (IChange) eventInfo.getDataObject();

			// Get Admin Session
			adminSession = repUtil.createAdminSession();
			formattedString = String.format("Re-loading objects with Admin Session: %s", adminSession.getCurrentUser());
			pxLogger.info(formattedString);
			amcObj = (IChange) adminSession.getObject(IChange.OBJECT_TYPE, amcObj.toString());
			formattedString = String.format("Re-loaded objects with Admin Session: %s", amcObj);
			pxLogger.info(formattedString);

			/*
			 * Validation for PX to be triggered when current status = Approval
			 * status
			 */

			simpleDateFormat = new SimpleDateFormat(constObj.amcredlinedateformat);
			String strTimeZone = constObj.amcredlinetimezone;
			simpleDateFormat.setTimeZone(TimeZone.getTimeZone(strTimeZone));

			// Fech the Affected Item table Iterator
			ITable affItemTable = amcObj.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
			Iterator<?> affItem = affItemTable.getTableIterator();
			Map<String, IItem> affItemMap = getaffectedItems(affItemTable);

			while (affItem.hasNext()) {
				String docSubType = "";
				String strDocSubType = "";
				IRow row = (IRow) affItem.next();
				IItem affitem = (IItem) row.getReferent();
				String oldLcp = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_LIFECYCLE_PHASE).toString();
				String newLcp = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_LIFECYCLE_PHASE).toString();
				String strItemType = affitem.getValue(ItemConstants.ATT_TITLE_BLOCK_ITEM_TYPE).toString();

				formattedString = String.format("Item type:%s", strItemType);
				pxLogger.info(formattedString);

				// Check for Doc Sub Type for Manufacturing Document and
				// Manufacturing Inspection Documents

				if (strMfgDoc.equalsIgnoreCase(strItemType) || strMfgInsp.equalsIgnoreCase(strItemType)) {

					formattedString = String.format("Document Item Number:%s",
							row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_NUMBER).toString());
					pxLogger.info(formattedString);

					if (null != service.getRedlinePage3Values(affitem, constObj.docSubTypeId)) {
						docSubType = service.getRedlinePage3Values(affitem, constObj.docSubTypeId);
						formattedString = String.format("Document Sub Type:%s", docSubType);
						pxLogger.info(formattedString);
						if (docSubType != null && (docSubType.equalsIgnoreCase("Line Setup Card")
								|| docSubType.equalsIgnoreCase("Packaging Visual Standard")
								|| docSubType.equalsIgnoreCase("IP Sheets"))) {

							strDocSubType = docSubType;
							pxLogger.info("Valid Document Sub Type:%s", strDocSubType);
						}
					}

				}

				formattedString = String.format("Item Old Lifecycle Phase:%s", oldLcp);
				pxLogger.info(formattedString);

				formattedString = String.format("Item New Lifecycle Phase:%s", newLcp);
				pxLogger.info(formattedString);

				// Based on Item Type of the Affected Item, call respective
				// methods to fetch the item details

				if (constObj.pvItemType.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside Production Version Item Type Loop");
					service.pvSheet(amcObj, affitem, oldLcp, newLcp, pvSheetMap, pvSheetNewList, pvSheetRevisedList,
							affItemMap);
					validItemList.add(affitem);
				} else if (constObj.mbomItemType.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside Mfg BOM Item Type Loop");
					strNewRev = null;
					if (null != row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString()) {
						strNewRev = row.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString();
					}
					service.bomSheet(affitem, oldLcp, strNewRev, bomSheetMap, bomSheetNewList, bomSheetRevisedList,
							affItemMap);
					validItemList.add(affitem);
				} else if (constObj.routingType.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside Mfg Routing Item Type Loop");
					service.routingSheet(affitem, oldLcp, routingsheetMap, routingSheetHeaderNewList,
							routingSheetHeaderRevisedList, affItemMap);
					validItemList.add(affitem);
				} else if (constObj.refOpItemType.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside Ref OP Item Type Loop");
					service.refOperationSheet(affitem, oldLcp, refOppSheetMap, refOppSheetNewList,
							refOppSheetRevisedList, affItemMap);
					validItemList.add(affitem);
				} else if (strLabel.equalsIgnoreCase(strItemType) || strAMIDMS.equalsIgnoreCase(strItemType)
						|| strDrawing.equalsIgnoreCase(strItemType) || strLabelDoc.equalsIgnoreCase(strItemType)
						|| strLabelFormat.equalsIgnoreCase(strItemType) || strManual.equalsIgnoreCase(strItemType)
						|| (strMfgDoc.equalsIgnoreCase(strItemType)
								&& strDocSubType.equalsIgnoreCase("Packaging Visual Standard"))
						|| (strMfgDoc.equalsIgnoreCase(strItemType)
								&& strDocSubType.equalsIgnoreCase("Line Setup Card"))
						|| (strMfgInsp.equalsIgnoreCase(strItemType) && strDocSubType.equalsIgnoreCase("IP Sheets"))) {
					pxLogger.info("Inside Documents Item Type Loop");
					service.getDocumentData(affitem, amcObj, oldLcp, row, strDocSubType, documentMap, newDocList,
							revDocList);
					validItemList.add(affitem);
				} else if (strRawMaterial.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside Raw Material Item Type loop");
					service.getRawMaterialData(amcObj, oldLcp, row, rawMatDetMap, newRMList, revRMList);
					validItemList.add(affitem);
				} else if (strComp.equalsIgnoreCase(strItemType) || strFGType.equalsIgnoreCase(strItemType)
						|| strPackMaterail.equalsIgnoreCase(strItemType)
						|| strDirForUse.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside Part Item Type loop");
					service.getPartDetailsData(amcObj, oldLcp, row, partDetailsMap, newPartList, revPartList);
					validItemList.add(affitem);
				} else if (strMfgStdTextKey.equalsIgnoreCase(strItemType)) {
					pxLogger.info("Inside std key Item Type Lopp");
					service.getStdTextKeyData(amcObj, oldLcp, row, stdKeyDetMap, newStdList, revStdList);
					validItemList.add(affitem);
				}

				// New Report file would be created if any Affected Item has
				// blank Lifecycle phase
				if (oldLcp.equalsIgnoreCase("")) {
					newFlag = true;
					strType = constObj.amcNewFile;

				}
				// Revised Report file would be created if any Affected Item has
				// Lifecycle phase other than blank
				else {
					revisedFlag = true;
					strType = constObj.amcRevFile;

				}
			} // while

			// To Identify whether New or Revised report files to be created
			if (newFlag == true && revisedFlag == true) {
				strType = constObj.amcBothFile;
			} else if (newFlag == true && revisedFlag == false) {
				strType = constObj.amcNewFile;
			} else if (newFlag == false && revisedFlag == true) {
				strType = constObj.amcRevFile;
			}

			formattedString = String.format("validItemList size:%s", validItemList.size());
			pxLogger.info(formattedString);

			if (!affItemTable.isEmpty() && !validItemList.isEmpty()) {
				changeDetailsList = service.getChangeData(amcObj, simpleDateFormat);
				CreateAMCExcelReport excelrep = new CreateAMCExcelReport();
				repFile = excelrep.generateReport(amcObj.getName(), changeDetailsList, documentMap, rawMatDetMap,
						partDetailsMap, pvSheetMap, bomSheetMap, routingsheetMap, refOppSheetMap, stdKeyDetMap,
						strType);

				// To attach the Redline Report to AMC Attachments tab if
				// Reports Files exist

				repUtil.attachReportToAMC(repFile, amcObj);
				actionResult = new ActionResult(ActionResult.STRING, constObj.redlinereportsuccess);
				if (repFile.isEmpty()) {
					actionResult = new ActionResult(ActionResult.STRING, constObj.redlinereportfailed);
				}
			} else {
				actionResult = new ActionResult(ActionResult.STRING, constObj.noAffItems);
				pxLogger.info(constObj.noAffItems);
			}
		} catch (Exception e) {
			try {
				String strEmailContent = AMCReportUtil.getLocalizedMessage(constObj.emailcontentsupport1,
						new Object[] { e }) + "\r\n" + constObj.emailcontentsupport2 + "\r\n"
						+ AMCReportUtil.exception2String(e);
				AMCReportUtil.mailAgent(constObj.agilesupport, constObj.emailfrom,
						AMCReportUtil.getLocalizedMessage(constObj.emailsubject, new Object[] { amcObj.toString() }),
						strEmailContent, constObj.smtphost);
			} catch (AddressException exception) {
				formattedString = String.format(constObj.mailnotsent, AMCReportUtil.exception2String(exception));
				pxLogger.info(formattedString);
			} catch (MessagingException ex) {
				formattedString = String.format(constObj.mailnotsent, AMCReportUtil.exception2String(ex));
				pxLogger.info(formattedString);
			}
			strMessage = constObj.redlinereportfailed;
			formattedString = String.format(": %s", strMessage);
			pxLogger.info(formattedString);
			StackTraceElement[] traces = e.getStackTrace();
			for (int inx = 0; inx < traces.length; inx++) {
				strMessage += traces[inx].toString() + "\n";
				pxLogger.error(traces[inx].toString());
			}
			return new EventActionResult(request, new ActionResult(ActionResult.EXCEPTION, new Exception(strMessage)));

		} finally {
			formattedString = String
					.format(": %s AMC Redline Report generation process is complete. Closing Session...", amcObj);
			pxLogger.info(formattedString);
			try {
				if (adminSession != null && adminSession.isOpen()) {
					adminSession.close();
				}
				pxLogger.info(
						"=========================================== Admin Session Closed. End of PX ============================================\n\n");
			} catch (APIException ex) {
				if (ex.getRootCause() != null) {
					strMessage = ex.getMessage() + " > " + ex.getRootCause().getMessage();
				} else {
					strMessage = ex.getMessage();
				}
				formattedString = String.format("Error while closing session:%s ", strMessage);
				pxLogger.info(formattedString);
				StackTraceElement[] traces = ex.getStackTrace();
				for (int inx = 0; inx < traces.length; inx++) {
					pxLogger.error(traces[inx].toString());
				}
			}
		}

		formattedString = String.format("final action result%s", actionResult);
		pxLogger.info(formattedString);

		return new EventActionResult(request, actionResult);
	}

	/**
	 * This method is for fetching the items of AMC
	 * 
	 * @param changeTable
	 *            ITable object
	 * @return newItemList : Map object
	 */

	public Map<String, IItem> getaffectedItems(ITable changeTable) {
		Map<String, IItem> itemsMap = new HashMap<>();
		Iterator<?> it;
		try {
			it = changeTable.getTableIterator();
			while (it.hasNext()) {
				IRow r = (IRow) it.next();
				IItem objItem = (IItem) r.getReferent();
				itemsMap.put(objItem.getName(), objItem);
			}
		} catch (Exception e) {
			formattedString = String.format("Error message: %s", e.getMessage());
			pxLogger.info(formattedString);
		}
		return itemsMap;
	}
}