/*
 * Copyright 2022 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
/* eslint-disable no-console */

import {
  loadCSS,
  toClassName,
  getMetadata,
  getExperimentConfig,
} from '../../scripts/scripts.js';

const percentformat = new Intl.NumberFormat('en-US', { style: 'percent', maximumSignificantDigits: 2 });
const countformat = new Intl.NumberFormat('en-US', { maximumSignificantDigits: 2 });
const significanceformat = {
  format: (value) => {
    if (value < 0.005) {
      return 'highly significant';
    } if (value < 0.05) {
      return 'significant';
    } if (value < 0.1) {
      return 'marginally significant';
    }
    return 'not significant';
  },
};
const bigcountformat = {
  format: (value) => {
    if (value > 1000000) {
      return `${countformat.format(value / 1000000)}M`;
    }
    if (value > 1000) {
      return `${countformat.format(value / 1000)}K`;
    }
    return countformat.format(value);
  },
};

/**
 * Create Badge if a Page is enlisted in a Helix Experiment
 * @return {Object} returns a badge or empty string
 */
async function createExperiment() {
  let selectedVariant;
  const engine = getMetadata('experimentation-engine') || 'franklin';
  if (engine === 'target') {
    const usp = new URLSearchParams(window.location.search);
    [, selectedVariant] = usp.get('experiment') ? usp.get('experiment').split('/') : [];
  } else {
    selectedVariant = (window.hlx && window.hlx.experiment && window.hlx.experiment.selectedVariant) ? window.hlx.experiment.selectedVariant : 'control';
  }
  const experiment = toClassName(getMetadata('experiment'));
  console.log('preview experiment', experiment);
  if (experiment) {
    let config;
    if (engine === 'target') {
      config = await getTargetExperimentConfig(experiment);
    } else {
      config = await getExperimentConfig(experiment);
    }

    const createVariant = (variantName) => {
      const variant = config.variants[variantName];
      const split = +variant.percentageSplit
        || 1 - config.variantNames.reduce((c, vn) => c + +config.variants[vn].percentageSplit, 0);
      const percentage = percentformat.format(split);
      const div = document.createElement('div');

      const experimentURL = new URL(window.location.href);
      // this will retain other query params such as ?rum=on
      experimentURL.searchParams.set('experiment', `${experiment}/${engine === 'target' ? Object.keys(config.variants).indexOf(variantName) + 1 : variantName}`);
      if (engine === 'target') {
        experimentURL.searchParams.set('token', 'yTcxjStWI3w5WBa6dFjUuZxhtLuBN4RrU0E8h4UVBzA');
      }

      if (engine === 'target') {
        div.className = `hlx-variant${config.variantNames.indexOf(variantName) === (Number(selectedVariant) - 1) ? ' hlx-variant-selected' : ' '}`;
      } else {
        div.className = `hlx-variant${selectedVariant === variantName ? ' hlx-variant-selected' : ' '}`;
      }
      div.innerHTML = `<div>
      <h5><code>${variantName}</code></h5>
        <p>${variant.label}</p>
        <p class="percentage">(${percentage} split)</p>
        <p class="performance"></p>
      </div>
      <div class="hlx-button"><a href="${experimentURL.href}">Simulate</a></div>`;
      return (div);
    };

    const manifestButton = config.manifest ? `<div class="hlx-button"><a href="${config.manifest}">Manifest</a></div>` : '';

    const div = document.createElement('div');
    div.className = 'hlx-experiment hlx-badge';
    div.classList.add(`hlx-experiment-status-${toClassName(config.status)}`);
    div.innerHTML = `Experiment: ${config.id} <span class="hlx-open"></span>
      <div class="hlx-popup hlx-hidden">
      <div class="hlx-popup-header">
        <div>
          <h4>${config.experimentName}</h4>
          <div class="hlx-details">${config.status}${config.audience ? ', ' : ''}${config.audience}${config.variants.control.blocks.length ? ', Blocks: ' : ''}${config.variants.control.blocks.join(',')}</div>
          <div class="hlx-info">How is it going?</div>
        </div>
        <div>
        ${manifestButton}
        </div>
      </div>
      <div class="hlx-variants"></div>
      </div>`;
    console.log(config.id);
    const popup = div.querySelector('.hlx-popup');

    const variantMap = {};

    div.addEventListener('click', () => {
      popup.classList.toggle('hlx-hidden');

      // the query is a bit slow, so I'm only fetching the results when the popup is opened
      const resultsURL = new URL('https://helix-pages.anywhere.run/helix-services/run-query@v2/rum-experiments');
      resultsURL.searchParams.set('experiment', experiment);
      if (window.hlx.sidekickConfig && window.hlx.sidekickConfig.host) {
        // restrict results to the production host, this also reduces query cost
        resultsURL.searchParams.set('domain', window.hlx.sidekickConfig.host);
      }
      fetch(resultsURL.href).then(async (response) => {
        const { results } = await response.json();

        const numberify = (obj) => Object.entries(obj).reduce((o, [k, v]) => {
          o[k] = Number.parseFloat(v);
          o[k] = Number.isNaN(o[k]) ? v : o[k];
          return o;
        }, {});

        const variantsAsNums = results.map(numberify);
        const totals = Object.entries(variantsAsNums.reduce((o, v) => {
          Object.entries(v).forEach(([k, val]) => {
            if (typeof val === 'number' && Number.isInteger(val) && k.startsWith('variant_')) {
              o[k] = (o[k] || 0) + val;
            } else if (typeof val === 'number' && Number.isInteger(val) && k.startsWith('control_')) {
              o[k] = val;
            }
          });
          return o;
        }, {})).reduce((o, [k, v]) => {
          o[k] = v;
          const vkey = k.replace(/^(variant|control)_/, 'variant_');
          const ckey = k.replace(/^(variant|control)_/, 'control_');
          const tkey = k.replace(/^(variant|control)_/, 'total_');
          if (o[ckey] && o[vkey]) {
            o[tkey] = o[ckey] + o[vkey];
          }
          return o;
        }, {});
        const richVariants = variantsAsNums.map((v) => ({
          ...v,
          allocation_rate: v.variant_experimentations / totals.total_experimentations,
        })).reduce((o, v) => {
          const variantName = v.variant;
          o[variantName] = v;
          return o;
        }, {
          control: {
            variant: 'control',
            ...Object.entries(variantsAsNums[0]).reduce((k, v) => {
              const [key, val] = v;
              if (key.startsWith('control_')) {
                k[key.replace(/^control_/, 'variant_')] = val;
              }
              return k;
            }, {}),
          },
        });
        const winner = variantsAsNums.reduce((w, v) => {
          if (v.variant_conversion_rate > w.conversion_rate && v.p_value < 0.05) {
            // eslint-disable-next-line no-param-reassign
            w.conversion_rate = v.variant_conversion_rate;
            // eslint-disable-next-line no-param-reassign
            w.p_value = v.p_value;
            // eslint-disable-next-line no-param-reassign
            w.variant = v.variant;
          }
          return w;
        }, {
          variant: 'control',
          p_value: Math.max(...variantsAsNums.map((v) => v.p_value)),
          conversion_rate: richVariants.control.variant_conversion_rate,
        });

        // add summary
        const summary = div.querySelector('.hlx-info');
        summary.innerHTML = `Showing results for ${bigcountformat.format(totals.total_experimentations)} visits and ${bigcountformat.format(totals.total_conversions)} conversions: `;
        if (totals.total_conversion_events < 500
          && winner.p_value > 0.05
          && variantsAsNums[0].remaining_runtime) {
          summary.innerHTML += ` not yet enough data to determine a winner. Keep going for about ${countformat.format(variantsAsNums[0].remaining_runtime)} days until you get ${bigcountformat.format((500 * totals.total_experimentations) / totals.total_conversion_events)} visits.`;
        } else if (totals.total_conversion_events < 500 && winner.p_value > 0.05) {
          summary.innerHTML += ` not yet enough data to determine a winner. Keep going until you get ${bigcountformat.format((500 * totals.total_experimentations) / totals.total_conversion_events)} visits.`;
        } else if (winner.p_value > 0.05) {
          summary.innerHTML += ' no significant difference between variants. In doubt, stick with <code>control</code>.';
        } else if (winner.variant === 'control') {
          summary.innerHTML += ' Stick with <code>control</code>. No variant is better than the control.';
        } else {
          summary.innerHTML += ` <code>${winner.variant}</code> is the winner.`;
        }

        // add traffic allocation to control and each variant
        Object.keys(variantMap).forEach((variantName) => {
          const variantDiv = variantMap[variantName];
          const percentage = variantDiv.querySelector('.percentage');
          percentage.innerHTML = `
            <span title="${countformat.format(richVariants[variantName].variant_conversion_events)} real events">${bigcountformat.format(richVariants[variantName].variant_conversions)} clicks</span> /
            <span title="${countformat.format(richVariants[variantName].variant_experimentation_events)} real events">${bigcountformat.format(richVariants[variantName].variant_experimentations)} visits</span>
            <span>(${percentformat.format(richVariants[variantName].variant_experimentations / totals.total_experimentations)} split)</span>
          `;
        });

        // add click rate and significance to each variant
        variantsAsNums.forEach((result) => {
          const variant = variantMap[result.variant];
          if (variant) {
            const performance = variant.querySelector('.performance');
            performance.innerHTML = `
              <span>click rate: ${percentformat.format(result.variant_conversion_rate)}</span>
              <span>vs. ${percentformat.format(result.control_conversion_rate)}</span>
              <span title="p value: ${result.p_value}" class="significance ${significanceformat.format(result.p_value).replace(/ /, '-')}">${significanceformat.format(result.p_value)}</span>
            `;
          }
        });
      });
    });

    const variants = div.querySelector('.hlx-variants');
    config.variantNames.forEach((vname) => {
      const variantDiv = createVariant(vname);
      variants.append(variantDiv);
      variantMap[vname] = variantDiv;
    });
    return (div);
  }
  return '';
}

/**
 * Decorates Preview mode badges and overlays
 * @return {Object} returns a badge or empty string
 */
async function decoratePreviewMode() {
  loadCSS('/tools/preview/preview.css');
  const overlay = document.createElement('div');
  overlay.className = 'hlx-preview-overlay';
  overlay.append(await createExperiment());
  document.body.append(overlay);
}

async function getTargetExperimentConfig(experimentId) {
  const targetExpConfig = {};
  const mapOffers = new Map();
  targetActivityJson?.options?.forEach((o) => mapOffers.set(o.optionLocalId, o.offerId));

  const experiences = targetActivityJson?.experiences?.map((exp) => {
    const experience = {};
    experience.percentage = exp.visitorPercentage;
    const offerId = mapOffers.get(exp.optionLocations[0].optionLocalId);
    let offerJson = '';
    if (offerId == 322724) {
      offerJson = offer322724;
    } else if (offerId == 322723) {
      offerJson = offer322723;
    }

    experience.url = offerJson.content.url;
    experience.offerName = offerJson.content.offerId;
    if (experience.offerName.toLowerCase().includes('control')) {
      experience.offerId = 'control';
    } else {
      experience.offerId = toClassName(experience.offerName);
    }
    return experience;
  });

  targetExpConfig.audience = 'Desktop';
  targetExpConfig.experimentName = targetActivityJson?.name;
  targetExpConfig.id = experimentId;
  targetExpConfig.manifest = '';
  targetExpConfig.status = targetActivityJson?.state;
  targetExpConfig.variantNames = experiences.map((exp) => exp.offerId);
  targetExpConfig.variants = {};
  experiences.forEach((exp) => {
    targetExpConfig.variants[exp.offerId] = {
      blocks: [''],
      label: exp.offerName,
      pages: [exp.url],
      percentageSplit: exp.percentage / 100,
    };
  });
  return targetExpConfig;
}

const targetActivityJson = JSON.parse('{\r\n    \"id\": 148141,\r\n    \"thirdPartyId\": \"ca5b2ab3-26e4-4d67-a0cc-8128876f0365\",\r\n    \"name\": \"FCBayern Newsletter Overlay A\/B Test\",\r\n    \"state\": \"approved\",\r\n    \"priority\": 0,\r\n    \"options\": [\r\n        {\r\n            \"optionLocalId\": 2,\r\n            \"name\": \"Offer2\",\r\n            \"offerId\": 322724,\r\n            \"offerTemplates\": []\r\n        },\r\n        {\r\n            \"optionLocalId\": 3,\r\n            \"name\": \"Offer3\",\r\n            \"offerId\": 322723,\r\n            \"offerTemplates\": []\r\n        }\r\n    ],\r\n    \"locations\": {\r\n        \"mboxes\": [\r\n            {\r\n                \"locationLocalId\": 0,\r\n                \"name\": \"target-newsletter-overlay\",\r\n                \"audienceIds\": []\r\n            }\r\n        ],\r\n        \"selectors\": []\r\n    },\r\n    \"experiences\": [\r\n        {\r\n            \"experienceLocalId\": 0,\r\n            \"name\": \"Experience A - Control\",\r\n            \"audienceIds\": [],\r\n            \"visitorPercentage\": 50,\r\n            \"optionLocations\": [\r\n                {\r\n                    \"locationLocalId\": 0,\r\n                    \"optionLocalId\": 2\r\n                }\r\n            ]\r\n        },\r\n        {\r\n            \"experienceLocalId\": 1,\r\n            \"name\": \"Experience B - Newsletter Overlay\",\r\n            \"audienceIds\": [],\r\n            \"visitorPercentage\": 50,\r\n            \"optionLocations\": [\r\n                {\r\n                    \"locationLocalId\": 0,\r\n                    \"optionLocalId\": 3\r\n                }\r\n            ]\r\n        }\r\n    ],\r\n    \"metrics\": [\r\n        {\r\n            \"metricLocalId\": 32767,\r\n            \"name\": \"MY PRIMARY GOAL\",\r\n            \"conversion\": true,\r\n            \"engagement\": \"score\",\r\n            \"action\": {\r\n                \"type\": \"count_once\"\r\n            },\r\n            \"mboxes\": [\r\n                {\r\n                    \"name\": \"target-newsletter-overlay\",\r\n                    \"successEvent\": \"mbox_shown\"\r\n                },\r\n                {\r\n                    \"name\": \"target-global-mbox\",\r\n                    \"successEvent\": \"mbox_shown\"\r\n                }\r\n            ],\r\n            \"clickTrackSelectors\": []\r\n        }\r\n    ],\r\n    \"reportingAudiences\": [],\r\n    \"workspace\": \"62749073\",\r\n    \"modifiedAt\": \"2022-10-28T16:06:38Z\"\r\n}');
const offer322724 = JSON.parse('{\r\n    \"id\": 322724,\r\n    \"name\": \"FCBayern - Newsletter Control\",\r\n    \"content\": {\r\n        \"offerId\": \"Target Newsletter Control\",\r\n        \"url\": \"https:\/\/main--fcbayern--hlxsites.hlx.live\/de\/\"\r\n    },\r\n    \"workspace\": \"62749073\",\r\n    \"modifiedAt\": \"2022-10-31T15:22:48Z\"\r\n}');
const offer322723 = JSON.parse('{\r\n  \"id\": 322723,\r\n  \"name\": \"FCBayern - Newsletter Overlay\",\r\n  \"content\": {\r\n      \"offerId\": \"Target Newsletter Overlay\",\r\n      \"url\": \"https:\/\/main--fcbayern--hlxsites.hlx.live\/experiments\/newsletter-overlay\/challenger-1\"\r\n  },\r\n  \"workspace\": \"62749073\",\r\n  \"modifiedAt\": \"2022-10-31T15:23:03Z\"\r\n}');

try {
  decoratePreviewMode();
} catch (e) {
  console.log(e);
}
