﻿using OfficeDevPnP.Core.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model.Teams
{
    /// <summary>
    /// Defines the Security settings for the Team
    /// </summary>
    public class TeamSecurity : BaseModel, IEquatable<TeamSecurity>
    {
        #region Public Members

        /// <summary>
        /// Defines the Owners of the Team
        /// </summary>
        public TeamSecurityUserCollection Owners { get; set; }

        /// <summary>
        /// Declares whether to clear existing owners before adding new ones
        /// </summary>
        public Boolean ClearExistingOwners { get; set; }

        /// <summary>
        /// Defines the Members of the Team
        /// </summary>
        public TeamSecurityUserCollection Members { get; set; }

        /// <summary>
        /// Declares whether to clear existing members before adding new ones
        /// </summary>
        public Boolean ClearExistingMembers { get; set; }

        #endregion

        #region Comparison code

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}",
                Owners.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                Members.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TeamSecurity class
        /// </summary>
        /// <param name="obj">Object that represents TeamSecurity</param>
        /// <returns>Checks whether object is TeamSecurity class</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TeamSecurity))
            {
                return (false);
            }
            return (Equals((TeamSecurity)obj));
        }

        /// <summary>
        /// Compares TeamSecurity object based on Owners, and Members
        /// </summary>
        /// <param name="other">TeamSecurity Class object</param>
        /// <returns>true if the TeamSecurity object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TeamSecurity other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Owners.DeepEquals(other.Owners) &&
                this.Members.DeepEquals(other.Members)
                );
        }

        #endregion
    }
    }
