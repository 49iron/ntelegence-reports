/**
 * Ntelegence Reports — Email Delivery via Resend
 * Sends .docx reports as email attachments.
 */

const { Resend } = require('resend');

let _resend = null;
function resend() {
  if (!_resend) _resend = new Resend(process.env.RESEND_API_KEY);
  return _resend;
}

/**
 * Send a report email with .docx attachment.
 *
 * @param {object} opts
 * @param {string|string[]} opts.to          - recipient email(s)
 * @param {string}          opts.subject     - email subject
 * @param {string}          opts.clientName  - e.g. "RNR Tire Midwest"
 * @param {string}          opts.period      - e.g. "May 31 – Jun 7, 2026"
 * @param {Buffer}          opts.docxBuffer  - the generated .docx as a Buffer
 * @param {string}          opts.filename    - attachment filename
 */
async function sendReport(opts) {
  const recipients = Array.isArray(opts.to) ? opts.to : [opts.to];

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#333">
      <div style="background:#1F3864;padding:20px 24px;border-radius:4px 4px 0 0">
        <img src="https://ntelegence.com/wp-content/uploads/ntelegence-logo-white.png"
             alt="Ntelegence" style="height:32px" />
      </div>
      <div style="padding:28px 24px;background:#f9f9f9;border:1px solid #e5e5e5;border-top:none">
        <h2 style="margin:0 0 8px;color:#1F3864">Call Intelligence Report Ready</h2>
        <p style="margin:0 0 16px;color:#555">
          Your weekly call intelligence executive summary for
          <strong>${opts.clientName}</strong> is attached.
        </p>
        <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px">
          <tr style="background:#e8f0f7">
            <td style="padding:8px 12px;font-weight:bold;color:#1F3864;width:40%">Report period</td>
            <td style="padding:8px 12px">${opts.period}</td>
          </tr>
          <tr style="background:#fff">
            <td style="padding:8px 12px;font-weight:bold;color:#1F3864">Prepared by</td>
            <td style="padding:8px 12px">Ntelegence Communication Intelligence</td>
          </tr>
          <tr style="background:#e8f0f7">
            <td style="padding:8px 12px;font-weight:bold;color:#1F3864">Contact</td>
            <td style="padding:8px 12px">
              <a href="mailto:jscherer@ntelegence.com" style="color:#2B6589">jscherer@ntelegence.com</a>
            </td>
          </tr>
        </table>
        <p style="margin:0 0 4px;font-size:13px;color:#888">
          Open the attached .docx in Microsoft Word or Google Docs.
        </p>
      </div>
      <div style="padding:12px 24px;background:#1F3864;border-radius:0 0 4px 4px;
                  font-size:12px;color:#aac;text-align:center">
        Ntelegence · No Dead Ends · ntelegence.com
      </div>
    </div>
  `;

  const result = await resend().emails.send({
    from:    'Ntelegence Reports <reports@ntelegence.com>',
    to:      recipients,
    subject: opts.subject,
    html,
    attachments: [{
      filename: opts.filename,
      content:  opts.docxBuffer.toString('base64'),
    }],
  });

  console.log(`✉  Sent "${opts.subject}" → ${recipients.join(', ')} [${result.data?.id ?? result.id}]`);
  return result;
}

module.exports = { sendReport };
