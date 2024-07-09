package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.*;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.collections4.IteratorUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

import static ser.Utils.loadTableRows;
import static ser.Utils.saveWBInboxExcel;

public class WBInboxMailLoad extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    IDocument mailTemplate = null;
    JSONObject projects = new JSONObject();
    ProcessHelper helper;

    @Override
    protected Object execute() {
        if (getBpm() == null)
            return resultError("Null BPM object");

        com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.WBInboxMailPaths.MainPath);



        try {
            helper = new ProcessHelper(Utils.session);
            JSONObject mcfg = Utils.getMailConfig();

            projects = Utils.getProjectWorkspaces(helper);

            mailTemplate = null;

            for(String prjn : projects.keySet()){
                IInformationObject prjt = (IInformationObject) projects.get(prjn);
                IDocument dtpl = Utils.getTemplateDocument(prjt, Conf.MailTemplates.Project);
                if(dtpl == null){continue;}
                mailTemplate = dtpl;
                break;
            }

            if(mailTemplate == null){throw new Exception("Mail template not found.");}

            List<IWorkbasket> wbs = Utils.bpm.getWorkbaskets();
            for (IWorkbasket wb : wbs){
                IWorkbasket swb = Utils.bpm.getWorkbasket(wb.getID());
                String dUserLogin = Utils.getUserIDfromWorkbasket(wb.getID());
                IUser user = getDocumentServer().getUserByLoginName(getSes(),dUserLogin);
                if(user!=null && user.getLicenseType() == LicenseType.TECHNICAL_USER) continue;
                if(Objects.equals(swb.getName(), "AgentService")){continue;}
                runWorkbasket(swb, mcfg, mailTemplate);
            }

            log.info("Tested.");
        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(),10);
        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
    private IInformationObject getProject(String prjn){
        if(projects.has(prjn)){return (IInformationObject) projects.get(prjn);}
        if(projects.has("!" + prjn)){return null;}

        IInformationObject iprj = Utils.getProjectWorkspace(prjn, helper);
        if(iprj == null){
            projects.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }
        projects.put(prjn, iprj);
        return iprj;
    }
    private void runWorkbasket(IWorkbasket swb, JSONObject mcfg, IDocument mtpl) throws Exception {
        String wbMail = swb.getNotifyEMail();

        log.info("WB : " + swb.getName());
        log.info(" *** mail : " + wbMail);
        log.info(" *** fullname : " + swb.getFullName());
        log.info(" *** accessible : " + swb.isAccessible());

        if(!swb.isAccessible()){return;}

        JSONObject prjDocs = new JSONObject();
        if(wbMail == null || wbMail.isEmpty()){return;}

        IWorkbasketContent wbco = swb.getWorkbasketContent();
        wbco.query();
        List<ITask> tasks = wbco.getTasks();
        if(tasks.size() < 1){return;}


        log.info("    -> start ");
        int tcnt = 0;
        JSONObject docs = new JSONObject();
        for(ITask task : tasks){

            tcnt++;
            log.info(" *** task [" + tcnt + "] : " + task.getDisplayName());

            String clid = task.getClassID();
            log.info(" *** clid [" + tcnt + "] : " + clid);
            if(!clid.equals(Conf.ClassIDs.Transmittal)
                    && !clid.equals(Conf.ClassIDs.SubReview)
                    && !clid.equals(Conf.ClassIDs.ReviewMain)){continue;}

            IProcessInstance proi = task.getProcessInstance();
            if(proi == null){continue;}
            log.info(" *** proi [" + tcnt + "] : " + proi.getDisplayName());

            IDocument wdoc = (IDocument) proi.getMainInformationObject();
            if(wdoc == null){continue;}
            log.info(" *** wdoc [" + tcnt + "] : " + wdoc.getDisplayName());

            String prjn = wdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn == null || prjn.isEmpty()){continue;}
            log.info(" *** prjn [" + tcnt + "] : " + prjn);

            if(docs.has(task.getID())){continue;}

            log.info("    -> class-name : " + task.getName());
            log.info("    -> class-id.doc : " + wdoc.getClassID());
            log.info("    -> class-id.task : " + clid);
            log.info("    -> display : " + wdoc.getDisplayName());

            docs.put(task.getID(), task);
        }


        if(docs == null || docs.length() < 1){return;}

        String uniqueId = UUID.randomUUID().toString();

        String mailExcelPath = Utils.exportDocument(mtpl, Conf.WBInboxMailPaths.MainPath, Conf.MailTemplates.Project + "[" + uniqueId + "]");
        List<String> dids = IteratorUtils.toList(docs.keys());

        loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, dids.size());

        JSONObject mbms = new JSONObject();
        mbms.put("Fullname", swb.getFullName());
        mbms.put("Count", dids.size() + "");

        int dcnt = 0;
        for(String zdid : dids){
            dcnt++;
            if(!docs.has(zdid)){continue;}
            ITask xtsk = (ITask) docs.get(zdid);
            IProcessInstance proi = xtsk.getProcessInstance();
            IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

            IWorkbasket orwb = xtsk.getOriginalWorkbasket();
            String owbn = "";
            if(orwb != null && orwb.getID() != swb.getID()){
                owbn = orwb.getFullName();
            }


            Date tbgn = null, tend = new Date();
            if(xtsk.getReadyDate() != null){
                tbgn = xtsk.getReadyDate();
            }

            long durd  = 0L;
            double durh  = 0.0;
            if(tend != null && tbgn != null) {
                long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
            }

            String rcvf = "", rcvo = "";
            if(xtsk.getPreviousWorkbasket() != null){
                rcvf = xtsk.getPreviousWorkbasket().getFullName();
            }
            if(tbgn != null){
                rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
            }

            String prjn = "",  mdno = "", mdrn = "", mdnm = "";
            if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            }
            if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
            }
            if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
            }
            if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
            }

            mbms.put("Task" + dcnt, xtsk.getName());
            mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
            mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
            mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
            mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
            mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
            mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
            mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
            mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
            mbms.put("DurDay" + dcnt, durd + "");
            mbms.put("DurHour" + dcnt, durh + "");
            mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
            mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
        }

        saveWBInboxExcel(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, mbms);

        String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath,
                Conf.WBInboxMailPaths.MainPath + "/" + Conf.MailTemplates.Project + "[" + uniqueId + "].html");
        JSONObject mail = new JSONObject();

        mail.put("To", wbMail);
        mail.put("Subject",
                "Workbox Notification ({Count}) Task is waiting your action"
                        .replace("{Count}", dids.size() + "")
        );
        //mail.put("Subject", "Reminder > " + prjn + " / " + swb.getFullName());
        mail.put("BodyHTMLFile", mailHtmlPath);

        try {
            //Utils.sendHTMLMail(mcfg, mail);
            Utils.sendHTMLMail(mail, null);
        }catch(Exception ex){
            log.error("EXCP [Send-Mail] : " + ex.getMessage());
        }
        log.info("    -> finish ");
    }

}