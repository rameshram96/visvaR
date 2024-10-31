.pkg_quotes <- c(
  "Statistics: The art of never having to say you're certain",
  "In God we trust; all others bring data. - W. Edwards Deming",
  "Statistical thinking will one day be as necessary for efficient citizenship as the ability to read and write. - H.G. Wells",
  "The plural of anecdote is not data.",
  "There are three kinds of lies: lies, damned lies, and statistics. - Benjamin Disraeli",
  "Statistics are like bikinis. What they reveal is suggestive, but what they conceal is vital. - Aaron Levenstein",
  "If your experiment needs statistics, you ought to have done a better experiment. - Ernest Rutherford",
  "I'm a statistician. I'm not good at math, but I can understand probability... probably.",
  "In data science, 80% of time spent is preparing data, 20% of time spent is complaining about preparing data.",
  "The combination of some data and an aching desire for an answer does not ensure that a reasonable answer can be extracted from a given body of data. - John Tukey",
  "The best thing about being a statistician is that you get to play in everyone's backyard.",
  "Correlation does not imply causation, but it does waggle its eyebrows suggestively and gesture furtively while mouthing 'look over there'.",
  "Correlation doesn't imply causation, but it does have a suggestive eyebrow wiggle.",
  "Garbage in, garbage out, but with statistics, garbage in, gospel out.",
  "The data may not contain the answer. The combination of some data and an aching desire for an answer does not ensure that a reasonable answer can be extracted.",
  "All models are wrong, but some are useful. - George Box",
  "The average human has one breast and one testicle. - Des McHale",
  "A model is a lie that helps you see the truth. - Howard Skipper",
  "No data is better than bad data, but bad data is often better than no data.",
  "Small samples are unreliable, large samples are impossible.",
  "The sample mean is always significant if your sample size is large enough.",
  "Statistics is the grammar of science. - Karl Pearson",
  "Without data, you're just another person with an opinion. - W. Edwards Deming",
  "Numbers don't lie, but liars use numbers.",
  "Data scientists are like modern day explorers - collecting bits instead of specimens.",
  "The world is one big data problem. - Andrew McAfee",
  "In data science, there's no such thing as a typical day.",
  "Probability is the most important concept in modern science, especially as nobody has the slightest notion what it means. - Bertrand Russell",
  "Random is not random.",
  "Probability is expectation founded upon partial knowledge. - George Boole",
  "The best we can hope for in this life is a reasonable level of significance.",
  "The only statistical significance that matters is practical significance.",
  "p < 0.05: The point at which scientists start believing in magic.",
  "If at first you don’t succeed, try two more times so that your failure is statistically significant",
  "If at first you don't succeed, try again. Then quit. No use being a damn fool about it. But at least you'll have n > 1",
  "Statisticians don't fail - they just achieve a high p-value",
  "I'm not failing, I'm collecting data points on what doesn't work",
  "Success is just statistically significant failure reduction",
  "My failures are normally distributed - all over the place",
  "Three statisticians walk into a bar... what are the odds? Pretty high actually, given our sample size",
  "I'm not procrastinating, I'm waiting for a larger sample size"
)

.onAttach <- function(libname, pkgname) {
  random_quote <- sample(.pkg_quotes, 1)
  msg <- paste0(
    "\n",
    "════════════════════════════════════════════════════════════════════\n",
    "Quote of the day: ", random_quote, "\n",
    "════════════════════════════════════════════════════════════════════\n"
  )

  packageStartupMessage(msg)
  packageStartupMessage("Thank you for installing 'visvaR'!")
  packageStartupMessage("For more information visit: https://rameshram96.github.io/visvaR/")
}

