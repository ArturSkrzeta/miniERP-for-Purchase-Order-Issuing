<h2>MiniERP for Purchase Order Issuing</h2>
<h3>Intro</h3>
<ul>
  <li>Script takes Purchase Request in excel format and converts it into Purchase Order printed into PDF file.</li>
  <li>PDF file is the final order file to be sent over to a supplier.</li>
  <li>In the meantime, script saves and tracks down the request details in the main table for the further analysis or audits.</li>
  <li>Both Putchase Request and Purchase Order gets archived at the end of macro's work.</li>
  <li>Script has the algorithm to generate a unique Purchase Order number (first tracker's column) so that it can be unique and doesn't duplicate previous ones.</li>
  <li>Also COMPANY prefix can be replaced by another ordering company daughter's name what is coded into algorithm as well. This allows to avoid worst case scenario where we generate two different orders under one number for the same company name.</li>
  <li>miniERP excel handles also back-end dabase of vendors and ordering company data that is being printed on Purchase Order.</li>
</ul>
<h3>Demo</h3>
<img src="images/tracker.JPG">
