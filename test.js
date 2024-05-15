
//---------------------------------------Some Tests---------------------------------------
//test1(); //Folder manangement tests
function test1() {
  matchingFolderPath = findMatchingFolder(proposalDirectory, "12345");
  console.log("testy testy testy");
  if (matchingFolderPath === null) {
    console.log("no match found. Test failed");
  } else {
    console.log("match found. Test passed");
    console.log("matchingFolderPath:", matchingFolderPath);
  }
  nullFolderPath = findMatchingFolder(proposalDirectory, "99999");
  if (nullFolderPath === null) {
    console.log("no match found. Test passed");
  } else {
    console.log("match found. Test failed");
    console.log("nullFolderPath:", nullFolderPath);
  }
  createProposalFolder("54321", "Test Title");
  console.log("54321 - Test Title folder attempted to be created");
}
//test2(); // Doc generation tests
async function test2() {
  try {
    const proposalNumber = "54321";
    const docContent = await getDocContent(proposalNumber);

    // Validate docContent before proceeding
    if (!Array.isArray(docContent) || docContent.length === 0) {
      console.error("docContent is empty or not an array");
      return; // Exit the function if validation fails
    }

    let matchingFolderPath = findMatchingFolder(
      proposalDirectory,
      proposalNumber,
    );

    if (matchingFolderPath) {
      console.log("folder found. Updating template...");
      updateTemplate(docContent, matchingFolderPath);
    } else {
      console.log("folder not found. Creating folder...");
      // Additional code for folder creation and subsequent actions
    }
  } catch (error) {
    // Log any errors that occur during the execution of getDocContent
    console.error("Error occurred in test2 function:", error);
  }
}

