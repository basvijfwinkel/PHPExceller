<?php
namespace PHPExceller;

use PHPExceller\CachedObjectStorage\PHPExceller_CachedObjectStorage_Memory;

/**
 * PHPExceller_CachedObjectStorageFactory
 *
 * Copyright (c) 2021 PHPExceller
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExceller
 * @package    PHPExceller_CachedObjectStorage
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_CachedObjectStorageFactory
{
    const cache_in_memory               = 'Memory';

    /**
     * Name of the method used for cell cacheing
     *
     * @var string
     */
    private static $cacheStorageMethod = null;

    /**
     * Name of the class used for cell cacheing
     *
     * @var string
     */
    private static $cacheStorageClass = null;

    /**
     * List of all possible cache storage methods
     *
     * @var string[]
     */
    private static $storageMethods = array(
        self::cache_in_memory
    );

    /**
     * Default arguments for each cache storage method
     *
     * @var array of mixed array
     */
    private static $storageMethodDefaultParameters = array(
        self::cache_in_memory => array(),
    );

    /**
     * Arguments for the active cache storage method
     *
     * @var array of mixed array
     */
    private static $storageMethodParameters = array();

    /**
     * Return the current cache storage method
     *
     * @return string|null
     **/
    public static function getCacheStorageMethod()
    {
        return self::$cacheStorageMethod;
    }

    /**
     * Return the current cache storage class
     *
     * @return PHPExceller_CachedObjectStorage_ICache|null
     **/
    public static function getCacheStorageClass()
    {
        return self::$cacheStorageClass;
    }

    /**
     * Return the list of all possible cache storage methods
     *
     * @return string[]
     **/
    public static function getAllCacheStorageMethods()
    {
        return self::$storageMethods;
    }

    /**
     * Return the list of all available cache storage methods
     *
     * @return string[]
     **/
    public static function getCacheStorageMethods()
    {
        $activeMethods = array();
        foreach (self::$storageMethods as $storageMethod)
        {
            $cacheStorageClass = "PHPExceller\\CachedObjectStorage\\".'PHPExceller_CachedObjectStorage_' . $storageMethod;
            if (call_user_func(array($cacheStorageClass, 'cacheMethodIsAvailable')))
            {
                $activeMethods[] = $storageMethod;
            }
        }
        return $activeMethods;
    }

    /**
     * Identify the cache storage method to use
     *
     * @param    string            $method        Name of the method to use for cell cacheing
     * @param    array of mixed    $arguments    Additional arguments to pass to the cell caching class
     *                                        when instantiating
     * @return boolean
     **/
    public static function initialize($method = self::cache_in_memory, $arguments = array())
    {
        if (!in_array($method, self::$storageMethods))
        {
            return false;
        }

        $cacheStorageClass = "PHPExceller\\CachedObjectStorage\\".'PHPExceller_CachedObjectStorage_'.$method;
        if (!call_user_func(array($cacheStorageClass, 'cacheMethodIsAvailable')))
        {
            return false;
        }

        self::$storageMethodParameters[$method] = self::$storageMethodDefaultParameters[$method];
        foreach ($arguments as $k => $v)
        {
            if (array_key_exists($k, self::$storageMethodParameters[$method]))
            {
                self::$storageMethodParameters[$method][$k] = $v;
            }
        }

        if (self::$cacheStorageMethod === null)
        {
            self::$cacheStorageClass = 'PHPExceller_CachedObjectStorage_' . $method;
            self::$cacheStorageMethod = $method;
        }
        return true;
    }

    /**
     * Initialise the cache storage
     *
     * @param    PHPExceller_Worksheet     $parent        Enable cell caching for this worksheet
     * @return    PHPExceller_CachedObjectStorage_ICache
     **/
    public static function getInstance(PHPExceller_Worksheet $parent)
    {
        $cacheMethodIsAvailable = true;
        if (self::$cacheStorageMethod === null)
        {
            $cacheMethodIsAvailable = self::initialize();
        }

        if ($cacheMethodIsAvailable)
        {
            $instance = new self::$cacheStorageClass(
                $parent,
                self::$storageMethodParameters[self::$cacheStorageMethod]
            );
            if ($instance !== null)
            {
                return $instance;
            }
        }

        return false;
    }

    /**
     * Clear the cache storage
     *
     **/
    public static function finalize()
    {
        self::$cacheStorageMethod = null;
        self::$cacheStorageClass = null;
        self::$storageMethodParameters = array();
    }
}
