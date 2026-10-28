/**
 * ============================================================
 * DSA MINI PROJECT: MOVIE RATING SORTER USING BUBBLE SORT
 * File: script.js
 * Purpose: Complete implementation of Bubble Sort algorithm
 *          applied to a real-world movie rating dataset.
 * Author: B.Tech Computer Science Student
 * ============================================================
 */

// -------------------------------------------------------------
// 1. SAMPLE DATASET (Initial Movies)
// -------------------------------------------------------------
const INITIAL_MOVIES = [
  { name: "Inception", rating: 8.8 },
  { name: "Interstellar", rating: 8.7 },
  { name: "3 Idiots", rating: 8.4 },
  { name: "The Dark Knight", rating: 9.0 },
  { name: "Avengers: Endgame", rating: 8.4 },
  { name: "Dangal", rating: 8.3 },
  { name: "KGF", rating: 8.2 },
  { name: "RRR", rating: 8.0 }
];

// Working copy of the movie list
let movies = JSON.parse(JSON.stringify(INITIAL_MOVIES));

// State trackers for sorting execution & statistics
let isSorting = false;
let comparisonsCount = 0;
let swapsCount = 0;
let passCount = 0;

// Speed delay in milliseconds (controlled by the slider)
let animationDelay = 600;

// -------------------------------------------------------------
// 2. DOM ELEMENT REFERENCES
// -------------------------------------------------------------
const movieContainer = document.getElementById("movie-container");
const statPass = document.getElementById("stat-pass");
const statComparisons = document.getElementById("stat-comparisons");
const statSwaps = document.getElementById("stat-swaps");
const statTotal = document.getElementById("stat-total");
const statusMessage = document.getElementById("status-message");
const speedRange = document.getElementById("speed-range");
const speedLabel = document.getElementById("speed-label");
const btnSortDesc = document.getElementById("btn-sort-desc");
const btnSortAsc = document.getElementById("btn-sort-asc");
const btnReset = document.getElementById("btn-reset");
const btnAdd = document.getElementById("btn-add");
const movieNameInput = document.getElementById("movie-name");
const movieRatingInput = document.getElementById("movie-rating");

// -------------------------------------------------------------
// 3. UTILITY HELPER: ASYNC SLEEP / PAUSE
// -------------------------------------------------------------
/**
 * Promisified sleep function to pause execution for a given duration.
 * This allows us to visualize each comparison and swap in the browser.
 * @param {number} ms - Milliseconds to pause
 */
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// -------------------------------------------------------------
// 4. RENDER FUNCTION
// -------------------------------------------------------------
/**
 * Renders the movies array as cards in the grid.
 * Optional highlights:
 *  - comparingIndices: [j, j + 1] -> colored in blue/cyan
 *  - swappingIndices: [j, j + 1] -> colored in red/coral
 *  - sortedIndices: Set of settled indices -> colored in emerald green
 */
function renderMovies(comparingIndices = [], swappingIndices = [], sortedIndices = new Set()) {
  movieContainer.innerHTML = "";

  movies.forEach((movie, index) => {
    const card = document.createElement("div");
    card.classList.add("movie-card");
    card.id = `movie-card-${index}`;

    // Apply visual highlight classes
    if (swappingIndices.includes(index)) {
      card.classList.add("swapping");
    } else if (comparingIndices.includes(index)) {
      card.classList.add("comparing");
    } else if (sortedIndices.has(index)) {
      card.classList.add("sorted");
    }

    // Calculate rating percentage for progress bar (rating out of 10)
    const percentage = Math.min(Math.max((movie.rating / 10) * 100, 0), 100);

    card.innerHTML = `
      <div class="movie-header-row">
        <span class="movie-rank">#${index + 1}</span>
        <button class="movie-delete-btn" title="Remove Movie" onclick="deleteMovie(${index})" ${isSorting ? 'disabled' : ''}>✕</button>
      </div>
      <div class="movie-name">${escapeHTML(movie.name)}</div>
      <div class="movie-rating-bar">
        <span class="rating-badge">⭐ ${movie.rating.toFixed(1)}</span>
        <span style="font-size: 0.8rem; color: #94a3b8;">/ 10</span>
      </div>
      <div class="rating-bar-visual">
        <div class="rating-bar-fill" style="width: ${percentage}%;"></div>
      </div>
    `;

    movieContainer.appendChild(card);
  });

  // Update total count
  statTotal.textContent = movies.length;
}

/**
 * Basic HTML escaping to prevent XSS in movie titles
 */
function escapeHTML(str) {
  return str.replace(/[&<>'"]/g, 
    tag => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      "'": '&#39;',
      '"': '&quot;'
    }[tag] || tag)
  );
}

// -------------------------------------------------------------
// 5. UPDATE STATISTICS IN UI
// -------------------------------------------------------------
function updateStats() {
  statPass.textContent = passCount;
  statComparisons.textContent = comparisonsCount;
  statSwaps.textContent = swapsCount;
}

// -------------------------------------------------------------
// 6. TOGGLE CONTROLS (ENABLE / DISABLE DURING SORTING)
// -------------------------------------------------------------
function setControlsDisabled(disabled) {
  isSorting = disabled;
  btnSortDesc.disabled = disabled;
  btnSortAsc.disabled = disabled;
  btnReset.disabled = disabled;
  btnAdd.disabled = disabled;
  movieNameInput.disabled = disabled;
  movieRatingInput.disabled = disabled;

  // Also disable/enable delete buttons
  const deleteBtns = document.querySelectorAll(".movie-delete-btn");
  deleteBtns.forEach(btn => btn.disabled = disabled);
}

// -------------------------------------------------------------
// 7. CORE BUBBLE SORT ALGORITHM WITH VISUALIZATION
// -------------------------------------------------------------
/**
 * Manual Bubble Sort Algorithm Implementation.
 * 
 * DSA Concept Breakdown:
 * - Compares adjacent elements: arr[j] and arr[j + 1].
 * - For Ascending (Lowest to Highest):
 *     if arr[j].rating > arr[j + 1].rating -> swap them
 * - For Descending (Highest to Lowest):
 *     if arr[j].rating < arr[j + 1].rating -> swap them
 * - Optimized with early exit:
 *     A boolean flag `swapped` tracks if any exchange occurred.
 *     If no elements were swapped during a pass, the list is already sorted!
 * 
 * @param {boolean} ascending - true for Lowest->Highest, false for Highest->Lowest
 */
async function bubbleSort(ascending = false) {
  const n = movies.length;

  if (n <= 1) {
    statusMessage.textContent = "The list has 1 or fewer movies. It is already sorted!";
    return;
  }

  // Lock UI buttons so user does not interrupt the algorithm
  setControlsDisabled(true);

  // Reset counters for a fresh run
  comparisonsCount = 0;
  swapsCount = 0;
  passCount = 0;
  updateStats();

  const sortDirectionText = ascending ? "Lowest → Highest (Ascending)" : "Highest → Lowest (Descending)";
  statusMessage.textContent = `Starting Bubble Sort in ${sortDirectionText} order...`;
  await sleep(animationDelay);

  // Set to track settled (sorted) element indices
  const sortedIndices = new Set();

  /**
   * OUTER LOOP: Controls the number of passes
   * Runs (n - 1) times in the worst case.
   */
  for (let i = 0; i < n - 1; i++) {
    passCount = i + 1;
    updateStats();

    // OPTIMIZATION FLAG:
    // Tracks if any swap happens in this pass.
    // If swapped remains false, array is sorted early (Best Case O(n)).
    let swapped = false;

    statusMessage.textContent = `Pass ${i + 1} started. Comparing adjacent pairs...`;
    await sleep(animationDelay * 0.7);

    /**
     * INNER LOOP: Compares adjacent elements
     * Notice: The upper bound is (n - i - 1) because the last 'i' elements
     * are already in their correct final positions from previous passes.
     */
    for (let j = 0; j < n - i - 1; j++) {
      comparisonsCount++;
      updateStats();

      const movieA = movies[j];
      const movieB = movies[j + 1];

      // Highlight the two adjacent elements being compared (Cyan / Blue)
      statusMessage.textContent = `Pass ${i + 1}: Comparing "${movieA.name}" (${movieA.rating}) with "${movieB.name}" (${movieB.rating})`;
      renderMovies([j, j + 1], [], sortedIndices);
      await sleep(animationDelay);

      /**
       * COMPARISON LOGIC:
       * - Descending: If current rating < next rating -> swap (bring higher rating forward)
       * - Ascending:  If current rating > next rating -> swap (bring lower rating forward)
       */
      let conditionToSwap = false;
      if (ascending) {
        conditionToSwap = movieA.rating > movieB.rating;
      } else {
        conditionToSwap = movieA.rating < movieB.rating;
      }

      if (conditionToSwap) {
        // Highlight the two cards in red/coral for swapping
        statusMessage.textContent = `Wrong order! Swapping "${movieA.name}" (${movieA.rating}) and "${movieB.name}" (${movieB.rating})...`;
        renderMovies([], [j, j + 1], sortedIndices);
        await sleep(animationDelay);

        /**
         * MANUAL SWAPPING LOGIC:
         * Using a standard temporary variable:
         *   temp = arr[j];
         *   arr[j] = arr[j + 1];
         *   arr[j + 1] = temp;
         */
        const temp = movies[j];
        movies[j] = movies[j + 1];
        movies[j + 1] = temp;

        swapsCount++;
        swapped = true;
        updateStats();

        // Render newly swapped positions
        renderMovies([], [j, j + 1], sortedIndices);
        await sleep(animationDelay * 0.7);
      } else {
        statusMessage.textContent = `Correct order: No swap needed for "${movieA.name}" and "${movieB.name}".`;
        await sleep(animationDelay * 0.5);
      }
    }

    // At the end of Pass i, the element at index (n - 1 - i) has bubbled to its final spot!
    sortedIndices.add(n - 1 - i);
    renderMovies([], [], sortedIndices);

    /**
     * EARLY EXIT OPTIMIZATION:
     * If no two elements were swapped by the inner loop, then the list is already sorted!
     */
    if (!swapped) {
      statusMessage.textContent = `⚡ Optimization Triggered! No swaps occurred in Pass ${i + 1}. The list is already fully sorted!`;
      // Mark all elements as sorted
      for (let k = 0; k < n; k++) {
        sortedIndices.add(k);
      }
      renderMovies([], [], sortedIndices);
      await sleep(animationDelay);
      break;
    }
  }

  // Mark all elements as settled
  for (let k = 0; k < n; k++) {
    sortedIndices.add(k);
  }
  renderMovies([], [], sortedIndices);

  statusMessage.textContent = `🎉 Sorting Complete! Total Passes: ${passCount}, Comparisons: ${comparisonsCount}, Swaps: ${swapsCount}.`;

  // Unlock UI
  setControlsDisabled(false);
}

// -------------------------------------------------------------
// 8. EVENT HANDLERS & USER INTERACTIONS
// -------------------------------------------------------------

/**
 * Initiates the Bubble Sort
 * @param {boolean} ascending - true for lowest to highest, false for highest to lowest
 */
function startBubbleSort(ascending) {
  if (isSorting) return;
  bubbleSort(ascending);
}

/**
 * Resets the movie list back to the initial sample data
 */
function resetMovieList() {
  if (isSorting) return;
  movies = JSON.parse(JSON.stringify(INITIAL_MOVIES));
  comparisonsCount = 0;
  swapsCount = 0;
  passCount = 0;
  updateStats();
  statusMessage.textContent = "Reset to default sample movies. Ready to sort!";
  renderMovies();
}

/**
 * Handles adding a new movie from the user input form
 */
function handleAddMovie() {
  if (isSorting) return;

  const nameVal = movieNameInput.value.trim();
  const ratingVal = parseFloat(movieRatingInput.value);

  // Input validation
  if (!nameVal) {
    alert("Please enter a valid movie name.");
    movieNameInput.focus();
    return;
  }

  if (isNaN(ratingVal) || ratingVal < 0 || ratingVal > 10) {
    alert("Please enter a rating between 0.0 and 10.0.");
    movieRatingInput.focus();
    return;
  }

  // Push new movie into array
  movies.push({
    name: nameVal,
    rating: parseFloat(ratingVal.toFixed(1))
  });

  // Clear input fields
  movieNameInput.value = "";
  movieRatingInput.value = "";
  movieNameInput.focus();

  statusMessage.textContent = `Added "${nameVal}" with rating ${ratingVal.toFixed(1)}. Ready to sort!`;
  renderMovies();
}

/**
 * Deletes a movie at a specific index
 * @param {number} index - Index of movie to remove
 */
function deleteMovie(index) {
  if (isSorting) return;
  const removed = movies.splice(index, 1);
  statusMessage.textContent = `Removed "${removed[0].name}".`;
  renderMovies();
}

/**
 * Speed slider change listener
 */
speedRange.addEventListener("input", (e) => {
  animationDelay = parseInt(e.target.value, 10);
  if (animationDelay <= 300) {
    speedLabel.textContent = `Fast (${animationDelay}ms)`;
  } else if (animationDelay <= 800) {
    speedLabel.textContent = `Normal (${animationDelay}ms)`;
  } else {
    speedLabel.textContent = `Slow (${animationDelay}ms)`;
  }
});

// -------------------------------------------------------------
// 9. INITIALIZATION ON PAGE LOAD
// -------------------------------------------------------------
window.addEventListener("DOMContentLoaded", () => {
  renderMovies();
  updateStats();
});
