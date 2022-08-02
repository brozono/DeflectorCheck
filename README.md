# DeflectorCheck

This tool will allow the review the usage of Deflectors for Egg Inc across identified contracts and group members.

At this time Windows only.

## Setup

Create a config file following the example: [Example.json](https://github.com/brozono/DeflectorCheck/blob/main/Examples/Example.json)

## Output

The output will be placed at your RoamingFolder:
C:\User\<username>\AppData\Roaming\DeflectorCheck\

There will be a folder for the Group with data from the coops and contracts to avoid going to the server for repeat runs
There will be a <GroupName>.xlsx file

## Running

Run the executable with the config file as parameter
DeflectorCheck.exe <Config.json>

## Interpreting Output

### Speadsheet Settings

There is a settings sheet with the ability to set the following parameters.

These parameters are used in the Summary sheet only.
- Personal Ratio Pass Threshold
  - Indicates the percentage of time a deflector was equipped
- Coop Ratio Pass Threshold
  - Indicates the percentage of time a deflector was equipped relative to others in your coop (note 100% would be the average)
- Contract Ratio Pass Threshold
  - Indicates the percentage of time a deflector was equipped relative to others in your group across all coops for that contract (note 100% would be the average)

This parameter is used on all sheets
- Max Contract Size Threshold
  - This is the size above which a contract will not be used in average calculations. To be used in case large contracts are outliers for deflector use.

### Summary Sheet

This sheet will show whether a group member was considered to have equipped a deflector based on meeting one or more of the following:
- Slotted at end of contract
- The personal percentage of time the deflector was equipped met or exceeded the personal threshold
- The coop percentage of time the deflector was equipped met or exceeded the coop threshold
- The contract percentage of time the deflector was equipped met or exceeded the contract threshold

## Slotted Sheet

This sheet will show whether a group member had a deflector slotted at the end of a contract.

Note
- If the contract is still active for the member and the deflector is equipped then this should show Yes
- If the contract has been completed and the member has exited the contract then the deflector must have been equipped at that time to show Yes

## Personal Sheet

This sheet shows the percentage of time that a member had a deflector equipped from the beginning to the estimated end of the contract. Estimated end of contract is based on token timer, coop laying rate, and typical group speed.
(will only apply if deflect was not unequippped before contract completed/exitted by member).

This sheet has conditional coloring
- Green means top 1/3 of the group's percentage for a given contract.
- No Color means middle 1/3 of the group's percentage for a given contract.
- Red means bottom 1/3 of the group's percentage for a given contract.

There are shades of each color to represent where in this group the member is. The more dark the green the closer to the top and the more dark the red the closer to the bottom.

## Coop Sheet

This sheet shows the percentage of time that a member had a deflector equipped relative to the other members in the same coop. A value of 100% is the average!

This sheet has conditional coloring
- Green means top 1/3 of the group's percentage for a given contract.
- No Color means middle 1/3 of the group's percentage for a given contract.
- Red means bottom 1/3 of the group's percentage for a given contract.

There are shades of each color to represent where in this group the member is. The more dark the green the closer to the top and the more dark the red the closer to the bottom.

## Contract Sheet

This sheet shows the percentage of time that a member had a deflector equipped relative to the other members in the group. A value of 100% is the average!

This sheet has conditional coloring
- Green means top 1/3 of the group's percentage for a given contract.
- No Color means middle 1/3 of the group's percentage for a given contract.
- Red means bottom 1/3 of the group's percentage for a given contract.

There are shades of each color to represent where in this group the member is. The more dark the green the closer to the top and the more dark the red the closer to the bottom.
