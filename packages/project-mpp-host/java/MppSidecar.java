import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import org.mpxj.ProjectFile;
import org.mpxj.Task;
import org.mpxj.mspdi.MSPDIWriter;
import org.mpxj.reader.UniversalProjectReader;

/**
 * PROJECT-018 — GenOffice MPP conversion sidecar (protocol v1).
 *
 * One-shot converter invoked by the host launcher as a DIRECT argument
 * array (never a shell string):
 *
 *   java -Djava.awt.headless=true -Xmx… -cp mpxj.jar:lib/* \
 *        MppSidecar.java <input.mpp> <output.mspdi> <requestId>
 *
 * Contract:
 *   - The MSPDI payload is written ONLY to the output file; stdout carries
 *     EXACTLY ONE JSON status line (the frame). Stdout can therefore never
 *     contaminate MSPDI parsing.
 *   - exit 0 + ok=true : conversion succeeded, output file written.
 *   - exit 0 + ok=false with error code MPP_UNSUPPORTED_FORMAT : the input
 *     is not a readable/recognized MPP project (unknown format, corrupt
 *     container, or an unreadable e.g. password-protected file).
 *   - exit != 0 : unexpected failure (IO error, crash) — the host reports
 *     MPP_SIDECAR_EXIT with the stderr tail; the frame, if any, was not
 *     produced.
 *
 * Determinism note: MSPDIWriter stamps <CurrentDate> with the current time
 * (a non-semantic save timestamp the accepted importer ignores entirely);
 * every data byte it writes is otherwise a pure in-memory transformation
 * (re-verified in the PROJECT-017 spike across two independent runs).
 *
 * This program performs no network access at the application level and
 * executes no code from the input file; it reads/writes only the two paths
 * it is given as arguments. Independently of that source-level posture,
 * the host launcher runs this program inside a kernel-enforced
 * network-isolated context (a Linux user + network namespace whose only
 * interface is a down loopback) under its default 'required' isolation
 * policy — and refuses to run it at all when the host cannot provide that
 * mechanism (fail closed).
 */
public class MppSidecar {

  public static void main(String[] args) {
    if (args.length != 3) {
      System.err.println("usage: MppSidecar <input.mpp> <output.mspdi> <requestId>");
      System.exit(2);
      return;
    }
    String inputPath = args[0];
    String outputPath = args[1];
    String requestId = args[2];

    ProjectFile project;
    try {
      project = new UniversalProjectReader().read(new File(inputPath));
      if (project == null) {
        frame(requestId, false, "MPP_UNSUPPORTED_FORMAT", "no project could be read from the input");
        return;
      }
    } catch (Throwable ex) {
      // Unrecognized/corrupt/password-protected containers all surface
      // here — deterministic refusal with the original reason preserved.
      frame(requestId, false, "MPP_UNSUPPORTED_FORMAT",
          "input could not be read as a supported project: " + rootMessage(ex));
      return;
    }

    try {
      int tasks = project.getTasks().size();
      int resources = project.getResources().size();
      int calendars = project.getCalendars().size();
      int predecessorLinks = 0;
      for (Task task : project.getTasks()) {
        predecessorLinks += task.getPredecessors().size();
      }
      int assignments = project.getResourceAssignments().size();

      // PROJECT-020 — the detected source format (the frame field the
      // protocol contract has always declared as "present on success";
      // the compatibility report's honest sourceVersion). For MPP inputs
      // this is the byte-true container generation ("MPP8"/"MPP9"/
      // "MPP12"/"MPP14"), NOT the filename or product provenance — e.g.
      // an MPP9-era file re-saved by a newer Project reports "MPP14".
      String fileType = project.getProjectProperties().getFileType();
      String format = "MPP".equals(fileType)
          ? "MPP" + project.getProjectProperties().getMppFileType()
          : fileType;

      new MSPDIWriter().write(project, new File(outputPath));

      StringBuilder sb = new StringBuilder(256);
      sb.append("{\"version\":1,\"requestId\":").append(jsonString(requestId));
      sb.append(",\"ok\":true");
      sb.append(",\"format\":").append(jsonString(format));
      sb.append(",\"counts\":{\"tasks\":").append(tasks);
      sb.append(",\"resources\":").append(resources);
      sb.append(",\"calendars\":").append(calendars);
      sb.append(",\"predecessorLinks\":").append(predecessorLinks);
      sb.append(",\"assignments\":").append(assignments);
      sb.append("}}");
      System.out.println(sb.toString());
    } catch (Throwable ex) {
      // Unexpected failure — crash loudly (nonzero exit); the host maps this
      // to MPP_SIDECAR_EXIT with the stderr tail. No partial output is
      // trusted: the host only reads the output file on frame ok=true.
      StringWriter sw = new StringWriter();
      ex.printStackTrace(new PrintWriter(sw, true));
      System.err.print(sw.toString());
      System.exit(1);
    }
  }

  /** Emit the failure frame (single JSON line on stdout, exit 0). */
  private static void frame(String requestId, boolean ok, String code, String message) {
    StringBuilder sb = new StringBuilder(256);
    sb.append("{\"version\":1,\"requestId\":").append(jsonString(requestId));
    sb.append(",\"ok\":false,\"error\":{\"code\":").append(jsonString(code));
    sb.append(",\"message\":").append(jsonString(message));
    sb.append("}}");
    System.out.println(sb.toString());
  }

  /** Minimal JSON string escaping (quotes, backslash, control chars). */
  private static String jsonString(String raw) {
    StringBuilder sb = new StringBuilder(raw.length() + 2);
    sb.append('"');
    for (int i = 0; i < raw.length(); i++) {
      char c = raw.charAt(i);
      switch (c) {
        case '"': sb.append("\\\""); break;
        case '\\': sb.append("\\\\"); break;
        case '\n': sb.append("\\n"); break;
        case '\r': sb.append("\\r"); break;
        case '\t': sb.append("\\t"); break;
        default:
          if (c < 0x20) {
            sb.append(String.format("\\u%04x", (int) c));
          } else {
            sb.append(c);
          }
      }
    }
    sb.append('"');
    return sb.toString();
  }

  /** Shortest useful exception message (root cause when chained). */
  private static String rootMessage(Throwable ex) {
    Throwable root = ex;
    while (root.getCause() != null && root.getCause() != root) {
      root = root.getCause();
    }
    String message = root.getMessage();
    String text = message == null ? root.getClass().getSimpleName() : message;
    return text.length() > 1000 ? text.substring(0, 1000) : text;
  }
}
