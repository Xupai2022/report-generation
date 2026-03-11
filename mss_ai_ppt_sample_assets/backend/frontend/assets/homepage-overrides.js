function replaceSectionNumber(heading, fromNumber, toNumber) {
  if (!heading) {
    return;
  }

  for (const node of heading.childNodes) {
    if (node.nodeType !== Node.TEXT_NODE || !node.textContent) {
      continue;
    }

    if (node.textContent.includes(fromNumber)) {
      node.textContent = node.textContent.replace(fromNumber, toNumber);
    }
  }
}

function updateHomeSectionNumbers() {
  const focusHeading = document.querySelector(".preconfig-left-col .preconfig-section:nth-of-type(2) h2");
  const templateHeading = document.querySelector(".preconfig-right-col .preconfig-template-section h2");

  replaceSectionNumber(focusHeading, "2.", "3.");
  replaceSectionNumber(templateHeading, "3.", "2.");
}

const observer = new MutationObserver(() => {
  updateHomeSectionNumbers();
});

observer.observe(document.documentElement, {
  childList: true,
  subtree: true,
});

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", updateHomeSectionNumbers, { once: true });
} else {
  updateHomeSectionNumbers();
}
